package doc

import (
	"encoding/binary"
	"errors"
	"unicode/utf8"
)

// listDef represents a list definition from the PlcfLst table.
type listDef struct {
	listID  uint32
	ordered bool   // true=ordered, false=unordered
	nfc     uint8  // number format code from LVLF (0=decimal, 1=upperRoman, 2=lowerRoman, 3=upperLetter, 4=lowerLetter, 23=bullet, etc.)
	lvlText string // level text template from xst (e.g. "(%1)" or "%1.")
}

// listOverride represents a list format override from the PlfLfo table.
type listOverride struct {
	listID uint32
}

// parsePlcfLst parses the list definition table (PlcfLst) from the Table stream.
// It extracts list IDs and attempts to determine ordered/unordered from LVLF nfc values.
//
// Per [MS-DOC] 2.8.12, PlfLst contains LSTF entries followed by LVLF entries.
// However, some DOC files report an lcb that only covers the LSTF array, with
// LVLF data stored immediately after in the table stream. This function handles
// both cases by extending the search range to the full table stream when LVLF
// data is not found within the lcb range.
func parsePlcfLst(tableData []byte, fc, lcb uint32) ([]listDef, error) {
	if lcb == 0 {
		return []listDef{}, nil
	}
	if uint64(fc)+uint64(lcb) > uint64(len(tableData)) {
		return nil, errors.New("PlcfLst data out of bounds")
	}

	data := tableData[fc : fc+lcb]
	if len(data) < 2 {
		return []listDef{}, nil
	}

	cLst := binary.LittleEndian.Uint16(data[0:2])
	// Each LSTF entry is 28 bytes, starting at offset 2
	lstfEnd := 2 + int(cLst)*28
	if lstfEnd > len(data) {
		return []listDef{}, nil
	}

	lists := make([]listDef, cLst)
	for i := 0; i < int(cLst); i++ {
		off := 2 + i*28
		lists[i].listID = binary.LittleEndian.Uint32(data[off : off+4])
		// Default to unordered
		lists[i].ordered = false
	}

	// Try to read LVLF entries after the LSTF array to determine ordered/unordered.
	// For each list, if fSimpleList is set there is 1 LVLF; otherwise 9 LVLFs.
	// We only inspect the first LVLF per list to check nfc (number format code).
	// nfc == 23 means bullet (unordered); anything else means ordered.
	//
	// LVLF data may be within the lcb range or immediately after it in the table
	// stream. Use the full table stream as the search boundary.
	lvlfAbsPos := int(fc) + lstfEnd
	for i := 0; i < int(cLst); i++ {
		flagsByte := data[2+i*28+26]
		fSimpleList := flagsByte & 0x01
		numLevels := 9
		if fSimpleList != 0 {
			numLevels = 1
		}

		for lvl := 0; lvl < numLevels; lvl++ {
			if lvlfAbsPos+28 > len(tableData) {
				break
			}
			// LVLF structure: offset 4 is nfc (uint8) - number format code
			nfc := tableData[lvlfAbsPos+4]
			if lvl == 0 {
				// nfc 23 = bullet (unordered), anything else = ordered
				lists[i].ordered = (nfc != 23)
				lists[i].nfc = nfc
			}
			// Skip LVLF (28 bytes) + variable data
			// cbGrpprlChpx at offset 24, cbGrpprlPapx at offset 25
			cbGrpprlChpx := int(tableData[lvlfAbsPos+24])
			cbGrpprlPapx := int(tableData[lvlfAbsPos+25])
			afterLvlf := lvlfAbsPos + 28
			afterGrpprl := afterLvlf + cbGrpprlPapx + cbGrpprlChpx
			// xst: uint16 length prefix + length*2 bytes of UTF-16 data
			if afterGrpprl+2 <= len(tableData) {
				xstLen := int(binary.LittleEndian.Uint16(tableData[afterGrpprl : afterGrpprl+2]))
				// Read the xst content for level 0 to get the lvlText template
				if lvl == 0 && xstLen > 0 && afterGrpprl+2+xstLen*2 <= len(tableData) {
					xstData := tableData[afterGrpprl+2 : afterGrpprl+2+xstLen*2]
					// Convert UTF-16LE to string, replacing placeholder bytes
					// In DOC xst, bytes 0x00-0x08 are level number placeholders
					// (0x00 = level 1, 0x01 = level 2, etc.)
					// In DOCX lvlText, these become %1, %2, etc.
					var lvlTextBuf []byte
					for k := 0; k+1 < len(xstData); k += 2 {
						ch := binary.LittleEndian.Uint16(xstData[k : k+2])
						if ch <= 8 {
							// Level placeholder: convert to %N format
							lvlTextBuf = append(lvlTextBuf, '%')
							lvlTextBuf = append(lvlTextBuf, byte('1'+ch))
						} else {
							// Regular character - append as UTF-8
							if ch < 128 {
								lvlTextBuf = append(lvlTextBuf, byte(ch))
							} else {
								// Encode as UTF-8
								r := rune(ch)
								buf := make([]byte, 4)
								n := utf8.EncodeRune(buf, r)
								lvlTextBuf = append(lvlTextBuf, buf[:n]...)
							}
						}
					}
					lists[i].lvlText = string(lvlTextBuf)
				}
				lvlfAbsPos = afterGrpprl + 2 + xstLen*2
			} else {
				lvlfAbsPos = len(tableData) // can't parse further
			}
		}
	}

	return lists, nil
}

// parsePlfLfo parses the list format override table (PlfLfo) from the Table stream.
// It extracts the list ID that each override references.
func parsePlfLfo(tableData []byte, fc, lcb uint32) ([]listOverride, error) {
	if lcb == 0 {
		return []listOverride{}, nil
	}
	if uint64(fc)+uint64(lcb) > uint64(len(tableData)) {
		return nil, errors.New("PlfLfo data out of bounds")
	}

	data := tableData[fc : fc+lcb]
	if len(data) < 4 {
		return []listOverride{}, nil
	}

	lfoMac := binary.LittleEndian.Uint32(data[0:4])

	overrides := make([]listOverride, 0, lfoMac)
	for i := uint32(0); i < lfoMac; i++ {
		off := 4 + i*16
		if off+4 > uint32(len(data)) {
			break
		}
		overrides = append(overrides, listOverride{
			listID: binary.LittleEndian.Uint32(data[off : off+4]),
		})
	}
	return overrides, nil
}
