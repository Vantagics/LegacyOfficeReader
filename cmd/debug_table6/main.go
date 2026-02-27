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

	fcPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x0102:])
	lcbPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x0106:])

	plcData := tableData[fcPlcfBtePapx : fcPlcfBtePapx+lcbPlcfBtePapx]
	n := (lcbPlcfBtePapx - 4) / 8

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
			if cb == 0 {
				if pos+1 < 512 {
					cb2 := int(page[pos+1])
					dataLen := cb2 * 2
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
				endPos := pos + 1 + dataLen
				if endPos > 512 {
					endPos = 512
				}
				if pos+3 < endPos {
					sprmData = page[pos+3 : endPos]
				}
			}

			// Parse sprms looking for sprmTDefTable with 2-byte length
			spos := 0
			for spos+2 <= len(sprmData) {
				opcode := binary.LittleEndian.Uint16(sprmData[spos:])
				spos += 2
				opSize := sprmOperandSize(opcode)
				if opSize == -1 {
					// Special case: sprmTDefTable and sprmTDefTable10 use 2-byte length
					if opcode == 0xD608 || opcode == 0xD605 {
						if spos+2 > len(sprmData) {
							break
						}
						opSize = int(binary.LittleEndian.Uint16(sprmData[spos:]))
						spos += 2
					} else {
						if spos >= len(sprmData) {
							break
						}
						opSize = int(sprmData[spos])
						spos++
					}
				}
				if spos+opSize > len(sprmData) {
					break
				}
				operand := sprmData[spos : spos+opSize]

				if opcode == 0xD608 {
					fmt.Printf("Page %d, run %d: sprmTDefTable (2-byte len)\n", pn, j)
					fmt.Printf("  Operand length: %d\n", len(operand))
					if len(operand) > 0 {
						itcMac := int(operand[0])
						fmt.Printf("  itcMac: %d\n", itcMac)
						if itcMac > 0 && len(operand) >= 1+2*(itcMac+1) {
							fmt.Printf("  rgdxaCenter:")
							for c := 0; c <= itcMac; c++ {
								val := int16(binary.LittleEndian.Uint16(operand[1+c*2:]))
								fmt.Printf(" %d", val)
							}
							fmt.Println()
							fmt.Printf("  Column widths:")
							for c := 0; c < itcMac; c++ {
								left := int16(binary.LittleEndian.Uint16(operand[1+c*2:]))
								right := int16(binary.LittleEndian.Uint16(operand[1+(c+1)*2:]))
								fmt.Printf(" %d", right-left)
							}
							fmt.Println()
						}
					}
				}
				spos += opSize
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
