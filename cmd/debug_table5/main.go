package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/doc"
)

func main() {
	// First, let's look at the raw PAPX data for table paragraphs
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Println("Error:", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Now let's look at the raw binary data
	adaptor, err := cfb.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Println("Error:", err)
		os.Exit(1)
	}
	defer adaptor.CloseFile()

	var wordDoc, table1, root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "WordDocument":
			wordDoc = dir
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

	tableReader, _ := adaptor.OpenObject(table1, root)
	tableSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tableData := make([]byte, tableSize)
	tableReader.Read(tableData)

	// Parse FIB
	fcPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x0102:])
	lcbPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x0106:])

	plcData := tableData[fcPlcfBtePapx : fcPlcfBtePapx+lcbPlcfBtePapx]
	n := (lcbPlcfBtePapx - 4) / 8

	// Scan all PAPX pages looking for table-related sprms
	for i := uint32(0); i < n; i++ {
		pnOffset := (n+1)*4 + i*4
		pn := binary.LittleEndian.Uint32(plcData[pnOffset:])
		pageOffset := pn * 512
		if pageOffset+512 > uint32(len(wordDocData)) {
			continue
		}
		page := wordDocData[pageOffset : pageOffset+512]
		crun := int(page[511])

		for j := 0; j < crun; j++ {
			bxPos := (crun+1)*4 + j*13
			if bxPos >= 511 {
				break
			}
			bOffset := int(page[bxPos])
			if bOffset == 0 {
				continue
			}
			pos := bOffset * 2
			if pos >= 512 {
				continue
			}
			cb := int(page[pos])
			var sprmData []byte
			var istd uint16
			if cb == 0 {
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

			// Scan for table-related sprms
			hasInTable := false
			hasRowEnd := false
			hasTDefTable := false
			spos := 0
			for spos+2 <= len(sprmData) {
				opcode := binary.LittleEndian.Uint16(sprmData[spos:])
				spos += 2
				opSize := sprmOperandSize(opcode)
				if opSize == -1 {
					if spos >= len(sprmData) {
						break
					}
					opSize = int(sprmData[spos])
					spos++
				}
				if spos+opSize > len(sprmData) {
					break
				}
				operand := sprmData[spos : spos+opSize]

				switch opcode {
				case 0x2416: // sprmPFInTable
					if len(operand) >= 1 && operand[0] != 0 {
						hasInTable = true
					}
				case 0x2417: // sprmPFTtp
					if len(operand) >= 1 && operand[0] != 0 {
						hasRowEnd = true
					}
				case 0xD608: // sprmTDefTable
					hasTDefTable = true
					if len(operand) > 0 {
						itcMac := int(operand[0])
						fmt.Printf("  sprmTDefTable: itcMac=%d, operand len=%d\n", itcMac, len(operand))
						if itcMac > 0 && len(operand) >= 1+2*(itcMac+1) {
							fmt.Printf("    rgdxaCenter:")
							for c := 0; c <= itcMac; c++ {
								val := int16(binary.LittleEndian.Uint16(operand[1+c*2:]))
								fmt.Printf(" %d", val)
							}
							fmt.Println()
						}
						// Dump raw
						fmt.Printf("    Raw: ")
						for k := 0; k < len(operand) && k < 40; k++ {
							fmt.Printf("%02x ", operand[k])
						}
						fmt.Println()
					}
				case 0xD605: // sprmTDefTable10
					fmt.Printf("  sprmTDefTable10 found\n")
				}
				spos += opSize
			}

			if hasInTable || hasRowEnd || hasTDefTable {
				fmt.Printf("Page %d, run %d: istd=%d inTable=%v rowEnd=%v tDefTable=%v\n",
					pn, j, istd, hasInTable, hasRowEnd, hasTDefTable)
				if hasTDefTable {
					fmt.Printf("  Full sprm data (%d bytes): ", len(sprmData))
					for k := 0; k < len(sprmData) && k < 80; k++ {
						fmt.Printf("%02x ", sprmData[k])
					}
					fmt.Println()
				}
			}
		}
	}
}

func sprmOperandSize(opcode uint16) int {
	spra := (opcode >> 13) & 0x07
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
		return -1
	case 7:
		return 3
	}
	return 0
}
