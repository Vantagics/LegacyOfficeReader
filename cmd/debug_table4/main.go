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

	// Parse FIB to get PlcfBtePapx location
	fcPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x0102:])
	lcbPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x0106:])
	fmt.Printf("PlcfBtePapx: fc=%d, lcb=%d\n", fcPlcfBtePapx, lcbPlcfBtePapx)

	plcData := tableData[fcPlcfBtePapx : fcPlcfBtePapx+lcbPlcfBtePapx]
	n := (lcbPlcfBtePapx - 4) / 8

	// Scan all PAPX pages looking for sprmTDefTable (0xD608)
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

			// Scan for sprmTDefTable (0xD608)
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

				if opcode == 0xD608 {
					fmt.Printf("\nFound sprmTDefTable at page %d, run %d\n", pn, j)
					fmt.Printf("  Operand length: %d bytes\n", len(operand))
					if len(operand) > 0 {
						itcMac := int(operand[0])
						fmt.Printf("  itcMac (num columns): %d\n", itcMac)
						if len(operand) >= 1+2*(itcMac+1) {
							fmt.Printf("  rgdxaCenter values:")
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
					// Dump raw hex
					fmt.Printf("  Raw hex: ")
					for k := 0; k < len(operand) && k < 60; k++ {
						fmt.Printf("%02x ", operand[k])
					}
					fmt.Println()
				}
				spos += opSize
			}
		}
	}
}

// sprmOperandSize returns the operand size for a given sprm opcode.
func sprmOperandSize(opcode uint16) int {
	sprmType := (opcode >> 13) & 0x07
	switch sprmType {
	case 0: // toggle
		return 1
	case 1: // byte
		return 1
	case 2: // word
		return 2
	case 3: // dword or long
		return 4
	case 4, 5: // short array
		return -1 // variable
	case 6: // variable
		return -1
	case 7: // 3 bytes
		return 3
	default:
		return -1
	}
}
