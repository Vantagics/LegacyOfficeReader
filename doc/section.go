package doc

import (
	"encoding/binary"
)

// sectionBreak represents a section break at a character position.
type sectionBreak struct {
	cpEnd   uint32 // CP of the last character in this section (the section mark)
	bkc     uint8  // break kind: 0=continuous, 1=new column, 2=new page, 3=even page, 4=odd page
	pgWidth uint16 // page width in twips (0 = use default)
	pgHeight uint16 // page height in twips (0 = use default)
}

// parsePlcfSed parses the Section Descriptor table (PlcfSed) from the Table stream.
// It returns section breaks with their CP positions and break types.
// wordDocData is needed to read the SEPX (Section Property Exception) data.
func parsePlcfSed(wordDocData, tableData []byte, fc, lcb uint32) []sectionBreak {
	if lcb == 0 {
		return nil
	}
	if uint64(fc)+uint64(lcb) > uint64(len(tableData)) {
		return nil
	}

	plcData := tableData[fc : fc+lcb]

	// PlcfSed: (n+1) CPs (uint32) + n SEDs (12 bytes each)
	// lcb = (n+1)*4 + n*12 = 4 + 16*n => n = (lcb - 4) / 16
	if lcb < 4 {
		return nil
	}
	n := (lcb - 4) / 16
	if n == 0 {
		return nil
	}

	var breaks []sectionBreak

	for i := uint32(0); i < n; i++ {
		// Read CP
		cpOff := i * 4
		if cpOff+4 > uint32(len(plcData)) {
			break
		}
		cp := binary.LittleEndian.Uint32(plcData[cpOff:])

		// Read SED entry (12 bytes): offset at (n+1)*4 + i*12
		sedOff := (n+1)*4 + i*12
		if sedOff+12 > uint32(len(plcData)) {
			break
		}

		// SED structure: 2 bytes fn, 4 bytes fcSepx, 2 bytes fnMpr, 4 bytes fcMpr
		fcSepx := binary.LittleEndian.Uint32(plcData[sedOff+2 : sedOff+6])

		sb := sectionBreak{cpEnd: cp, bkc: 2} // default: new page

		// Parse SEPX if valid
		if fcSepx != 0 && fcSepx != 0xFFFFFFFF && uint64(fcSepx)+2 <= uint64(len(wordDocData)) {
			cbSepx := binary.LittleEndian.Uint16(wordDocData[fcSepx:])
			sepxStart := fcSepx + 2
			if uint64(sepxStart)+uint64(cbSepx) <= uint64(len(wordDocData)) {
				sepxData := wordDocData[sepxStart : sepxStart+uint32(cbSepx)]
				parseSepxSprms(sepxData, &sb)
			}
		}

		breaks = append(breaks, sb)
	}

	return breaks
}

// parseSepxSprms parses section property sprms from SEPX data.
func parseSepxSprms(data []byte, sb *sectionBreak) {
	pos := 0
	for pos+2 <= len(data) {
		opcode := binary.LittleEndian.Uint16(data[pos:])
		pos += 2

		opSize := sprmOperandSize(opcode)
		if opSize == -1 {
			if pos >= len(data) {
				break
			}
			opSize = int(data[pos])
			pos++
		}
		if pos+opSize > len(data) {
			break
		}

		operand := data[pos : pos+opSize]
		pos += opSize

		switch opcode {
		case 0x3009: // sprmSBkc - section break kind (1 byte)
			if len(operand) >= 1 {
				sb.bkc = operand[0]
			}
		case 0xB01F: // sprmSXaPage - page width (uint16)
			if len(operand) >= 2 {
				sb.pgWidth = binary.LittleEndian.Uint16(operand)
			}
		case 0xB020: // sprmSYaPage - page height (uint16)
			if len(operand) >= 2 {
				sb.pgHeight = binary.LittleEndian.Uint16(operand)
			}
		}
	}
}
