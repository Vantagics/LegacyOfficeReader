package doc

import (
	"encoding/binary"
	"errors"
)

// paraFormatRun represents a paragraph's formatting properties over a character range.
type paraFormatRun struct {
	cpStart         uint32
	cpEnd           uint32
	props           ParagraphFormatting
	istd            uint16
	inTable         bool
	tableRowEnd     bool
	ilfo            uint16
	ilvl            uint8
	pageBreakBefore bool
	outLvl          uint8    // heading outline level (0-9, where 9 means body text)
	cellWidths      []int32  // table cell widths in twips (from sprmTDefTable)
}

// parsePapxSprms parses a byte sequence of paragraph property sprm entries
// and returns the resulting paraFormatRun.
func parsePapxSprms(data []byte) paraFormatRun {
	var pf paraFormatRun
	pf.outLvl = 9 // default: body text (no heading)
	pos := 0

	for pos+2 <= len(data) {
		opcode := binary.LittleEndian.Uint16(data[pos:])
		pos += 2

		// Determine operand size using shared helper from chpx.go
		opSize := sprmOperandSize(opcode)
		if opSize == -1 {
			// Variable length sprms.
			// sprmTDefTable (0xD608) uses a 2-byte (uint16) length prefix
			// per [MS-DOC] 2.6.4. All other variable-length sprms use a
			// 1-byte length prefix.
			if opcode == 0xD608 {
				if pos+2 > len(data) {
					break
				}
				opSize = int(binary.LittleEndian.Uint16(data[pos:]))
				pos += 2
			} else {
				if pos >= len(data) {
					break
				}
				opSize = int(data[pos])
				pos++ // skip the size byte
			}
		}

		if pos+opSize > len(data) {
			break
		}

		operand := data[pos : pos+opSize]
		pos += opSize

		switch opcode {
		case 0x2461: // sprmPJc - paragraph alignment (1 byte)
			if len(operand) >= 1 {
				pf.props.Alignment = operand[0]
				pf.props.AlignmentSet = true
			}
		case 0x2403: // sprmPJc80 - paragraph alignment, older version (1 byte)
			if len(operand) >= 1 {
				// Only set if sprmPJc (0x2461) hasn't been seen yet.
				// sprmPJc takes precedence over sprmPJc80.
				if !pf.props.AlignmentSet {
					pf.props.Alignment = operand[0]
					pf.props.AlignmentSet = true
				}
			}
		case 0x845E: // sprmPDxaLeft - left indent (int16)
			if len(operand) >= 2 {
				pf.props.IndentLeft = int32(int16(binary.LittleEndian.Uint16(operand)))
			}
		case 0x845D: // sprmPDxaRight - right indent (int16)
			if len(operand) >= 2 {
				pf.props.IndentRight = int32(int16(binary.LittleEndian.Uint16(operand)))
			}
		case 0x8460: // sprmPDxaLeft1 - first line indent (int16)
			if len(operand) >= 2 {
				pf.props.IndentFirst = int32(int16(binary.LittleEndian.Uint16(operand)))
			}
		case 0xA413: // sprmPDyaBefore - space before (uint16)
			if len(operand) >= 2 {
				pf.props.SpaceBefore = binary.LittleEndian.Uint16(operand)
			}
		case 0xA414: // sprmPDyaAfter - space after (uint16)
			if len(operand) >= 2 {
				pf.props.SpaceAfter = binary.LittleEndian.Uint16(operand)
			}
		case 0x6412: // sprmPDyaLine - line spacing (4 bytes: int16 dyaLine + int16 fMultLinespace)
			if len(operand) >= 4 {
				dyaLine := int16(binary.LittleEndian.Uint16(operand[0:2]))
				fMult := binary.LittleEndian.Uint16(operand[2:4])
				pf.props.LineSpacing = int32(dyaLine)
				// fMult=1 means multiple (value is in 240ths of a line)
				// fMult=0 and dyaLine>0 means "at least" (value in twips)
				// fMult=0 and dyaLine<0 means "exact" (absolute value in twips)
				if fMult == 1 {
					pf.props.LineRule = 0 // auto/multiple
				} else if dyaLine < 0 {
					pf.props.LineSpacing = int32(-dyaLine)
					pf.props.LineRule = 2 // exact
				} else {
					pf.props.LineRule = 1 // atLeast
				}
			}
		case 0x4600: // sprmPIstd - paragraph style index (uint16)
			if len(operand) >= 2 {
				pf.istd = binary.LittleEndian.Uint16(operand)
			}
		case 0x2416: // sprmPFInTable - in table flag (1 byte)
			if len(operand) >= 1 {
				pf.inTable = operand[0] != 0
			}
		case 0x2417: // sprmPFTtp - table row end flag (1 byte)
			if len(operand) >= 1 {
				pf.tableRowEnd = operand[0] != 0
			}
		case 0x460B: // sprmPIlfo - list override index (uint16)
			if len(operand) >= 2 {
				pf.ilfo = binary.LittleEndian.Uint16(operand)
			}
		case 0x260A: // sprmPIlvl - list level (1 byte)
			if len(operand) >= 1 {
				pf.ilvl = operand[0]
			}
		case 0x2407: // sprmPFPageBreakBefore - page break before (1 byte)
			if len(operand) >= 1 {
				pf.pageBreakBefore = operand[0] != 0
			}
		case 0x2640: // sprmPOutLvl - outline level (1 byte)
			if len(operand) >= 1 {
				pf.outLvl = operand[0]
			}
		case 0xD608: // sprmTDefTable - table column definitions (variable length)
			// Format: first byte = number of columns (itcMac)
			// Then (itcMac+1) int16 values for column boundary positions (rgdxaCenter)
			// Then itcMac TC structures (20 bytes each) - we skip these
			if len(operand) >= 1 {
				itcMac := int(operand[0])
				if itcMac > 0 && len(operand) >= 1+2*(itcMac+1) {
					widths := make([]int32, itcMac)
					for c := 0; c < itcMac; c++ {
						left := int16(binary.LittleEndian.Uint16(operand[1+c*2:]))
						right := int16(binary.LittleEndian.Uint16(operand[1+(c+1)*2:]))
						widths[c] = int32(right - left)
					}
					pf.cellWidths = widths
				}
			}
		case 0xD605: // sprmTDefTable10 - older table column definitions
			if len(operand) >= 1 {
				itcMac := int(operand[0])
				if itcMac > 0 && len(operand) >= 1+2*(itcMac+1) {
					widths := make([]int32, itcMac)
					for c := 0; c < itcMac; c++ {
						left := int16(binary.LittleEndian.Uint16(operand[1+c*2:]))
						right := int16(binary.LittleEndian.Uint16(operand[1+(c+1)*2:]))
						widths[c] = int32(right - left)
					}
					pf.cellWidths = widths
				}
			}
		}
		// Unknown sprms are silently skipped (operand already consumed above)
	}

	return pf
}

// parsePlcBtePapx parses the paragraph property exception table (PlcBtePapx)
// from the Table stream and returns paragraph format runs.
// The pieces parameter is used to convert FC byte offsets to CP character positions.
func parsePlcBtePapx(wordDocData, tableData []byte, fc, lcb uint32, pieces []piece) ([]paraFormatRun, error) {
	if lcb == 0 {
		return []paraFormatRun{}, nil
	}

	if uint64(fc)+uint64(lcb) > uint64(len(tableData)) {
		return nil, errors.New("PlcBtePapx data out of bounds")
	}

	plcData := tableData[fc : fc+lcb]

	// PlcBtePapx structure:
	// (n+1) FC values (uint32 each) followed by n PnBtePapx entries (4 bytes each)
	// lcb = (n+1)*4 + n*4 = 4 + 8*n  =>  n = (lcb - 4) / 8
	if lcb < 4 {
		return []paraFormatRun{}, nil
	}
	n := (lcb - 4) / 8
	if n == 0 {
		return []paraFormatRun{}, nil
	}

	var runs []paraFormatRun

	for i := uint32(0); i < n; i++ {
		// Read PnBtePapx entry: page number
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

		// PapxFkp: last byte is crun (count of runs)
		crun := int(page[511])
		if crun == 0 {
			continue
		}

		// FKP contains (crun+1) FC values (uint32 each) at the start
		fcArraySize := (crun + 1) * 4

		// After the FC array, there are crun BX entries (13 bytes each):
		// first byte is bOffset, remaining 12 bytes are PHE data (ignored)
		bxStart := fcArraySize

		for j := 0; j < crun; j++ {
			// Read FC range (these are byte offsets in the WordDocument stream)
			fcStartVal := binary.LittleEndian.Uint32(page[j*4:])
			fcEndVal := binary.LittleEndian.Uint32(page[(j+1)*4:])

			// Read BX entry (13 bytes): first byte is bOffset
			bxPos := bxStart + j*13
			if bxPos >= 511 {
				break
			}
			bOffset := int(page[bxPos])

			var pf paraFormatRun
			pf.outLvl = 9 // default

			if bOffset != 0 {
				pos := bOffset * 2
				if pos < 512 {
					cb := int(page[pos])
					var istd uint16
					var sprmData []byte

					if cb == 0 {
						// cb == 0: read next byte as cb2
						if pos+1 < 512 {
							cb2 := int(page[pos+1])
							dataLen := cb2 * 2
							if pos+4 <= 512 {
								istd = binary.LittleEndian.Uint16(page[pos+2 : pos+4])
							}
							endPos := pos + 2 + dataLen
							if endPos > 512 {
								endPos = 512
							}
							if pos+4 < endPos {
								sprmData = page[pos+4 : endPos]
							}
						}
					} else {
						// cb != 0: data is cb*2 - 1 bytes starting at pos+1
						dataLen := cb*2 - 1
						if dataLen >= 2 && pos+1+2 <= 512 {
							istd = binary.LittleEndian.Uint16(page[pos+1 : pos+3])
						}
						endPos := pos + 1 + dataLen
						if endPos > 512 {
							endPos = 512
						}
						if pos+3 < endPos {
							sprmData = page[pos+3 : endPos]
						}
					}

					if len(sprmData) > 0 {
						pf = parsePapxSprms(sprmData)
					}
					pf.istd = istd
				}
			}

			// Convert FC byte offsets to CP character positions using piece table
			cpStart, cpEnd := convertFCRangesToCP(fcStartVal, fcEndVal, pieces)
			pf.cpStart = cpStart
			pf.cpEnd = cpEnd

			runs = append(runs, pf)
		}
	}

	return runs, nil
}
