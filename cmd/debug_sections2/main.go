package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
)

func main() {
	adaptor, err := cfb.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer adaptor.CloseFile()

	var wordDoc, table0, table1, root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "WordDocument":
			wordDoc = dir
		case "0Table":
			table0 = dir
		case "1Table":
			table1 = dir
		case "Root Entry":
			root = dir
		}
	}

	wordDocReader, _ := adaptor.OpenObject(wordDoc, root)
	wordDocSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wordDocData := make([]byte, wordDocSize)
	wordDocReader.Read(wordDocData)

	flags := binary.LittleEndian.Uint16(wordDocData[0x0A:])
	fWhichTblStm := (flags >> 9) & 1
	var tableDir *cfb.Directory
	if fWhichTblStm == 1 {
		tableDir = table1
	} else {
		tableDir = table0
	}

	tableReader, _ := adaptor.OpenObject(tableDir, root)
	tableSize := binary.LittleEndian.Uint32(tableDir.StreamSize[:])
	tableData := make([]byte, tableSize)
	tableReader.Read(tableData)

	// Navigate FIB
	offset := 0x20
	csw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2

	readFcLcb := func(index int) uint32 {
		if int(cbRgFcLcb) <= index {
			return 0
		}
		off := offset + index*4
		if off+4 > len(wordDocData) {
			return 0
		}
		return binary.LittleEndian.Uint32(wordDocData[off:])
	}

	fcPlcfSed := readFcLcb(12)
	lcbPlcfSed := readFcLcb(13)
	fmt.Printf("fcPlcfSed=%d lcbPlcfSed=%d\n", fcPlcfSed, lcbPlcfSed)

	if lcbPlcfSed == 0 {
		fmt.Println("No section descriptors")
		return
	}

	plcData := tableData[fcPlcfSed : fcPlcfSed+lcbPlcfSed]
	n := (lcbPlcfSed - 4) / 16
	fmt.Printf("Number of sections: %d\n", n)

	for i := uint32(0); i < n; i++ {
		cp := binary.LittleEndian.Uint32(plcData[i*4:])
		sedOff := (n+1)*4 + i*12
		fcSepx := binary.LittleEndian.Uint32(plcData[sedOff+2 : sedOff+6])

		fmt.Printf("\nSection[%d]: cpEnd=%d fcSepx=0x%08X\n", i, cp, fcSepx)

		if fcSepx != 0 && fcSepx != 0xFFFFFFFF && uint64(fcSepx)+2 <= uint64(len(wordDocData)) {
			cbSepx := binary.LittleEndian.Uint16(wordDocData[fcSepx:])
			sepxStart := fcSepx + 2
			if uint64(sepxStart)+uint64(cbSepx) <= uint64(len(wordDocData)) {
				sepxData := wordDocData[sepxStart : sepxStart+uint32(cbSepx)]
				fmt.Printf("  SEPX: %d bytes\n", cbSepx)
				// Parse sprms
				pos := 0
				for pos+2 <= len(sepxData) {
					opcode := binary.LittleEndian.Uint16(sepxData[pos:])
					pos += 2
					opSize := sprmOperandSize(opcode)
					if opSize == -1 {
						if pos >= len(sepxData) {
							break
						}
						opSize = int(sepxData[pos])
						pos++
					}
					if pos+opSize > len(sepxData) {
						break
					}
					operand := sepxData[pos : pos+opSize]
					pos += opSize

					switch opcode {
					case 0x3009: // sprmSBkc
						fmt.Printf("  sprmSBkc (break kind): %d\n", operand[0])
					case 0xB01F: // sprmSXaPage
						w := binary.LittleEndian.Uint16(operand)
						fmt.Printf("  sprmSXaPage (page width): %d twips (%.1f in)\n", w, float64(w)/1440)
					case 0xB020: // sprmSYaPage
						h := binary.LittleEndian.Uint16(operand)
						fmt.Printf("  sprmSYaPage (page height): %d twips (%.1f in)\n", h, float64(h)/1440)
					case 0xB021: // sprmSDxaLeft (left margin)
						m := binary.LittleEndian.Uint16(operand)
						fmt.Printf("  sprmSDxaLeft (left margin): %d twips (%.1f in)\n", m, float64(m)/1440)
					case 0x9023: // sprmSDyaTop (top margin)
						m := int16(binary.LittleEndian.Uint16(operand))
						fmt.Printf("  sprmSDyaTop (top margin): %d twips (%.1f in)\n", m, float64(m)/1440)
					case 0xB022: // sprmSDxaRight (right margin)
						m := binary.LittleEndian.Uint16(operand)
						fmt.Printf("  sprmSDxaRight (right margin): %d twips (%.1f in)\n", m, float64(m)/1440)
					case 0x9024: // sprmSDyaBottom (bottom margin)
						m := int16(binary.LittleEndian.Uint16(operand))
						fmt.Printf("  sprmSDyaBottom (bottom margin): %d twips (%.1f in)\n", m, float64(m)/1440)
					case 0xB017: // sprmSDyaHdrTop (header distance)
						m := binary.LittleEndian.Uint16(operand)
						fmt.Printf("  sprmSDyaHdrTop (header dist): %d twips\n", m)
					case 0xB018: // sprmSDyaHdrBottom (footer distance)
						m := binary.LittleEndian.Uint16(operand)
						fmt.Printf("  sprmSDyaHdrBottom (footer dist): %d twips\n", m)
					case 0x3014: // sprmSDmBinFirst
						fmt.Printf("  sprmSDmBinFirst: %d\n", operand[0])
					case 0x3015: // sprmSDmBinOther
						fmt.Printf("  sprmSDmBinOther: %d\n", operand[0])
					case 0x500B: // sprmSCcolumns
						fmt.Printf("  sprmSCcolumns: %d\n", binary.LittleEndian.Uint16(operand))
					case 0x3228: // sprmSFTitlePage
						fmt.Printf("  sprmSFTitlePage (different first page): %d\n", operand[0])
					default:
						fmt.Printf("  sprm 0x%04X: %d bytes\n", opcode, opSize)
					}
				}
			}
		}
	}
}

func sprmOperandSize(opcode uint16) int {
	sgc := (opcode >> 13) & 0x07
	spra := (opcode >> 10) & 0x07
	_ = sgc
	switch spra {
	case 0:
		return 1
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
		return -1 // variable
	case 7:
		return 3
	default:
		return 1
	}
}
