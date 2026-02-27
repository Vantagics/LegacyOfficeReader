package doc

import (
	"encoding/binary"
	"errors"
	"fmt"
)

// charFormatRun represents a range of text with uniform character formatting.
type charFormatRun struct {
	cpStart uint32
	cpEnd   uint32
	props   CharacterFormatting
}

// icoToHex maps the legacy color index (sprmCIco) to a 6-digit hex RGB string.
var icoToHex = [...]string{
	"",       // 0 = auto / no color
	"000000", // 1 = black
	"0000FF", // 2 = blue
	"00FFFF", // 3 = cyan
	"00FF00", // 4 = green
	"FF00FF", // 5 = magenta
	"FF0000", // 6 = red
	"FFFF00", // 7 = yellow
	"FFFFFF", // 8 = white
}

// sprmOperandSize returns the operand byte count for a given sprm opcode
// based on the spra field (bits 13-15).
func sprmOperandSize(opcode uint16) int {
	spra := (opcode >> 13) & 0x07
	switch spra {
	case 0:
		return 1 // toggle
	case 1:
		return 1
	case 2:
		return 2
	case 3:
		return 4
	case 4:
		return 2
	case 5:
		return 2
	case 6:
		return -1 // variable: first byte is size
	case 7:
		return 3
	}
	return 0
}

// parseChpxSprms parses a byte sequence of character property sprm entries
// and returns the resulting CharacterFormatting.
func parseChpxSprms(data []byte, styles []styleDef, fonts []string) CharacterFormatting {
	var cf CharacterFormatting
	pos := 0

	for pos+2 <= len(data) {
		opcode := binary.LittleEndian.Uint16(data[pos:])
		pos += 2

		// Determine operand size
		opSize := sprmOperandSize(opcode)
		if opSize == -1 {
			// Variable length: first byte is the size of the operand
			if pos >= len(data) {
				break
			}
			opSize = int(data[pos])
			pos++ // skip the size byte
		}

		if pos+opSize > len(data) {
			break
		}

		operand := data[pos : pos+opSize]
		pos += opSize

		switch opcode {
		case 0x4A4F: // sprmCRgFtc0 - font index (uint16)
			if len(operand) >= 2 {
				fontIdx := int(binary.LittleEndian.Uint16(operand))
				if fonts != nil && fontIdx < len(fonts) && fonts[fontIdx] != "" {
					cf.FontName = fonts[fontIdx]
				} else {
					cf.FontName = fmt.Sprintf("font%d", fontIdx)
				}
			}
		case 0x4A43: // sprmCHps - font size in half-points (uint16)
			if len(operand) >= 2 {
				cf.FontSize = binary.LittleEndian.Uint16(operand)
			}
		case 0x0835: // sprmCFBold - toggle (1 byte)
			if len(operand) >= 1 {
				cf.Bold = operand[0] != 0
			}
		case 0x0836: // sprmCFItalic - toggle (1 byte)
			if len(operand) >= 1 {
				cf.Italic = operand[0] != 0
			}
		case 0x2A3E: // sprmCKul - underline type (1 byte)
			if len(operand) >= 1 {
				cf.Underline = operand[0]
			}
		case 0x2A42: // sprmCIco - color index (1 byte)
			if len(operand) >= 1 {
				idx := operand[0]
				if int(idx) < len(icoToHex) {
					cf.Color = icoToHex[idx]
				} else {
					cf.Color = "000000"
				}
			}
		case 0x6870: // sprmCCv - direct RGB color (COLORREF: 4 bytes, little-endian 0xBBGGRR)
			if len(operand) >= 3 {
				// COLORREF stores as [R, G, B, 0] in little-endian memory
				cf.Color = fmt.Sprintf("%02X%02X%02X", operand[0], operand[1], operand[2])
			}
		case 0x4A30: // sprmCIstd - character style index (uint16)
			if len(operand) >= 2 {
				cf.IstdChar = binary.LittleEndian.Uint16(operand)
			}
		case 0x6A03: // sprmCPicLocation - picture location in Data stream (int32)
			if len(operand) >= 4 {
				cf.PicLocation = int32(binary.LittleEndian.Uint32(operand))
				cf.HasPicLocation = true
			}
		}
		// Unknown sprms are silently skipped (operand already consumed above)
	}

	return cf
}

// parsePlcBteChpx parses the character property exception table (PlcBteChpx)
// from the Table stream and returns character format runs.
// The pieces parameter is used to convert FC byte offsets to CP character positions.
func parsePlcBteChpx(wordDocData, tableData []byte, fc, lcb uint32, styles []styleDef, fonts []string, pieces []piece) ([]charFormatRun, error) {
	if lcb == 0 {
		return []charFormatRun{}, nil
	}

	if uint64(fc)+uint64(lcb) > uint64(len(tableData)) {
		return nil, errors.New("PlcBteChpx data out of bounds")
	}

	plcData := tableData[fc : fc+lcb]

	// PlcBteChpx structure:
	// (n+1) FC values (uint32 each) followed by n PnBteChpx entries (4 bytes each)
	// lcb = (n+1)*4 + n*4 = 4 + 8*n  =>  n = (lcb - 4) / 8
	if lcb < 4 {
		return []charFormatRun{}, nil
	}
	n := (lcb - 4) / 8
	if n == 0 {
		return []charFormatRun{}, nil
	}

	var runs []charFormatRun

	for i := uint32(0); i < n; i++ {
		// Read PnBteChpx entry: page number
		pnOffset := (n+1)*4 + i*4
		if pnOffset+4 > uint32(len(plcData)) {
			break
		}
		pn := binary.LittleEndian.Uint32(plcData[pnOffset:])

		// The actual byte offset in WordDocument stream is pn * 512
		pageOffset := pn * 512
		if pageOffset+512 > uint32(len(wordDocData)) {
			continue // skip pages that exceed wordDocData
		}

		page := wordDocData[pageOffset : pageOffset+512]

		// ChpxFkp: last byte is crun (count of runs)
		crun := int(page[511])
		if crun == 0 {
			continue
		}

		// FKP contains (crun+1) FC values (uint32 each) at the start
		fcArraySize := (crun + 1) * 4
		if fcArraySize > 511 {
			continue
		}

		// After the FC array, there are crun 1-byte offsets (rgb)
		rgbStart := fcArraySize

		for j := 0; j < crun; j++ {
			// Read FC range (these are byte offsets in the WordDocument stream)
			fcStartVal := binary.LittleEndian.Uint32(page[j*4:])
			fcEndVal := binary.LittleEndian.Uint32(page[(j+1)*4:])

			// Read rgb offset
			if rgbStart+j >= 511 {
				break
			}
			rgb := int(page[rgbStart+j])

			var props CharacterFormatting
			if rgb != 0 {
				// rgb * 2 gives position within the FKP page
				chpxPos := rgb * 2
				if chpxPos < 512 {
					// First byte is cb (size of Chpx data)
					cb := int(page[chpxPos])
					if chpxPos+1+cb <= 512 {
						sprmData := page[chpxPos+1 : chpxPos+1+cb]
						props = parseChpxSprms(sprmData, styles, fonts)
					}
				}
			}

			// Convert FC byte offsets to CP character positions using piece table
			cpStart, cpEnd := convertFCRangesToCP(fcStartVal, fcEndVal, pieces)

			runs = append(runs, charFormatRun{
				cpStart: cpStart,
				cpEnd:   cpEnd,
				props:   props,
			})
		}
	}

	return runs, nil
}
