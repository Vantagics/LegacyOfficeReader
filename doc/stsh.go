package doc

import (
	"encoding/binary"
	"errors"
	"unicode/utf16"
)

// styleType represents the kind of a style definition.
type styleType uint8

const (
	styleTypeParagraph styleType = 1
	styleTypeCharacter styleType = 2
)

// styleDef represents a single style definition from the STSH.
type styleDef struct {
	name      string
	sti       uint16    // built-in style index (0-based)
	styleType styleType
	istdBase  uint16 // base style index; 0xFFF means no base
	charProps *CharacterFormatting
	paraProps *ParagraphFormatting
}

// parseSTSH parses the Style Sheet (STSH) from the Table stream.
// tableData is the full Table stream, fc is the offset and lcb is the length
// of the STSH data within that stream.
// fonts is the font name table (from SttbfFfn) used to resolve font indices in UPX data.
func parseSTSH(tableData []byte, fc, lcb uint32, fonts []string) ([]styleDef, error) {
	if lcb == 0 {
		return []styleDef{}, nil
	}

	if uint64(fc)+uint64(lcb) > uint64(len(tableData)) {
		return nil, errors.New("STSH data out of bounds")
	}

	data := tableData[fc : fc+lcb]

	// STSH starts with Stshi header.
	// Bytes 0-1: cbStshi (uint16) - size of Stshi fixed portion in bytes.
	if len(data) < 4 {
		return []styleDef{}, nil
	}
	cbStshi := binary.LittleEndian.Uint16(data[0:2])

	// Within Stshi (after cbStshi field):
	// Bytes 2-3: cstd (uint16) - count of styles
	// Bytes 4-5: cbSTDBaseInFile (uint16) - base size of each STD entry
	if int(cbStshi)+2 < 6 {
		// Not enough header data for cstd and cbSTDBaseInFile
		return []styleDef{}, nil
	}
	cstd := binary.LittleEndian.Uint16(data[2:4])
	cbSTDBaseInFile := binary.LittleEndian.Uint16(data[4:6])

	// STD entries start after the Stshi header.
	// The Stshi header occupies cbStshi + 2 bytes (2 for the cbStshi field itself).
	pos := int(cbStshi) + 2
	styles := make([]styleDef, 0, int(cstd))

	for i := 0; i < int(cstd); i++ {
		if pos+2 > len(data) {
			// No more data; fill remaining with empty styles
			for j := i; j < int(cstd); j++ {
				styles = append(styles, styleDef{})
			}
			break
		}

		// Each STD starts with cbStd (uint16) - size of this STD's data
		cbStd := binary.LittleEndian.Uint16(data[pos : pos+2])
		pos += 2

		if cbStd == 0 {
			// Empty/deleted style
			styles = append(styles, styleDef{})
			continue
		}

		if pos+int(cbStd) > len(data) {
			// Truncated STD; add empty and skip
			styles = append(styles, styleDef{})
			pos += int(cbStd)
			break
		}

		stdData := data[pos : pos+int(cbStd)]
		pos += int(cbStd)

		sd := parseSTD(stdData, cbSTDBaseInFile, fonts)
		styles = append(styles, sd)
	}

	return styles, nil
}

// parseSTD parses a single STD (Style Definition) entry.
func parseSTD(stdData []byte, cbSTDBaseInFile uint16, fonts []string) styleDef {
	sd := styleDef{}

	if len(stdData) < 4 {
		return sd
	}

	// Per [MS-DOC] 2.9.260 Stdf:
	// Word 0 (bytes 0-1): bits 0-11 = sti, bit 12 = fScratch, bit 13 = fInvalHeight,
	//                      bit 14 = fHasUpe, bit 15 = fMassCopy
	// Word 1 (bytes 2-3): bits 0-3 = stk (style kind), bits 4-15 = istdBase
	word0 := binary.LittleEndian.Uint16(stdData[0:2])
	sd.sti = word0 & 0x0FFF

	word1 := binary.LittleEndian.Uint16(stdData[2:4])
	stk := word1 & 0x000F
	sd.istdBase = (word1 >> 4) & 0x0FFF

	switch stk {
	case 1:
		sd.styleType = styleTypeParagraph
	case 2:
		sd.styleType = styleTypeCharacter
	default:
		sd.styleType = styleType(stk)
	}

	// After the fixed base (cbSTDBaseInFile bytes), parse the style name.
	nameOffset := int(cbSTDBaseInFile)
	if nameOffset+2 > len(stdData) {
		return sd
	}

	// Style name: uint16 length (in characters), followed by UTF-16LE encoded name,
	// followed by 2 null bytes (terminator).
	nameLen := binary.LittleEndian.Uint16(stdData[nameOffset : nameOffset+2])
	nameOffset += 2

	if nameLen > 0 && nameOffset+int(nameLen)*2 <= len(stdData) {
		u16 := make([]uint16, nameLen)
		for j := 0; j < int(nameLen); j++ {
			u16[j] = binary.LittleEndian.Uint16(stdData[nameOffset+j*2 : nameOffset+j*2+2])
		}
		sd.name = string(utf16.Decode(u16))
	}

	// Skip past the name + null terminator to find UPX data
	upxOffset := nameOffset + int(nameLen)*2 + 2 // +2 for null terminator
	// Align to even boundary
	if upxOffset%2 != 0 {
		upxOffset++
	}

	// Parse UPX (Universal Property eXception) data for style properties
	// For paragraph styles: first UPX is paragraph properties, second is character properties
	// For character styles: only one UPX for character properties
	if sd.styleType == styleTypeParagraph && upxOffset < len(stdData) {
		// First UPX: paragraph properties (2-byte istd + sprm data)
		if upxOffset+2 <= len(stdData) {
			cbUpx := int(binary.LittleEndian.Uint16(stdData[upxOffset:]))
			upxOffset += 2
			if cbUpx > 2 && upxOffset+cbUpx <= len(stdData) {
				// First 2 bytes are istd, rest is sprm data
				sprmData := stdData[upxOffset+2 : upxOffset+cbUpx]
				pf := parsePapxSprms(sprmData)
				sd.paraProps = &pf.props
			}
			if cbUpx > 0 && upxOffset+cbUpx <= len(stdData) {
				upxOffset += cbUpx
			}
			// Align to even boundary
			if upxOffset%2 != 0 {
				upxOffset++
			}
		}
		// Second UPX: character properties
		if upxOffset+2 <= len(stdData) {
			cbUpx := int(binary.LittleEndian.Uint16(stdData[upxOffset:]))
			upxOffset += 2
			if cbUpx > 0 && upxOffset+cbUpx <= len(stdData) {
				upxData := stdData[upxOffset : upxOffset+cbUpx]
				cf := parseChpxSprms(upxData, nil, fonts)
				sd.charProps = &cf
			}
		}
	} else if sd.styleType == styleTypeCharacter && upxOffset < len(stdData) {
		// Single UPX: character properties
		if upxOffset+2 <= len(stdData) {
			cbUpx := int(binary.LittleEndian.Uint16(stdData[upxOffset:]))
			upxOffset += 2
			if cbUpx > 0 && upxOffset+cbUpx <= len(stdData) {
				upxData := stdData[upxOffset : upxOffset+cbUpx]
				cf := parseChpxSprms(upxData, nil, fonts)
				sd.charProps = &cf
			}
		}
	}

	return sd
}
