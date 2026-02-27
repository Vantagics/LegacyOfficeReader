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
		if pn != 31 {
			continue
		}

		pageOffset := pn * 512
		page := wordDocData[pageOffset : pageOffset+512]
		crun := int(page[511])

		j := 0
		bxPos := (crun+1)*4 + j*13
		bOffset := int(page[bxPos])
		pos := bOffset * 2
		cb := int(page[pos])
		
		dataLen := cb*2 - 1
		endPos := pos + 1 + dataLen
		if endPos > 512 {
			endPos = 512
		}
		sprmData := page[pos+3 : endPos]
		
		// Simulate the fixed parsePapxSprms with 2-byte length for TDefTable
		spos := 0
		for spos+2 <= len(sprmData) {
			opcode := binary.LittleEndian.Uint16(sprmData[spos:])
			spos += 2
			
			spra := (opcode >> 13) & 0x07
			var opSize int
			
			switch spra {
			case 0, 1:
				opSize = 1
			case 2:
				opSize = 2
			case 3:
				opSize = 4
			case 4, 5:
				opSize = 2
			case 6:
				if opcode == 0xD608 {
					if spos+2 > len(sprmData) {
						fmt.Println("Not enough data for 2-byte length")
						goto done
					}
					opSize = int(binary.LittleEndian.Uint16(sprmData[spos:]))
					fmt.Printf("opcode=0x%04X: 2-byte length = %d (bytes: %02x %02x)\n", opcode, opSize, sprmData[spos], sprmData[spos+1])
					spos += 2
				} else {
					if spos >= len(sprmData) {
						goto done
					}
					opSize = int(sprmData[spos])
					spos++
				}
			case 7:
				opSize = 3
			}
			
			if spos+opSize > len(sprmData) {
				fmt.Printf("opcode=0x%04X: opSize=%d exceeds remaining data (%d)\n", opcode, opSize, len(sprmData)-spos)
				break
			}
			
			if opcode == 0xD608 {
				operand := sprmData[spos : spos+opSize]
				fmt.Printf("sprmTDefTable operand (first 30 bytes): ")
				for k := 0; k < 30 && k < len(operand); k++ {
					fmt.Printf("%02x ", operand[k])
				}
				fmt.Println()
				if len(operand) > 0 {
					itcMac := int(operand[0])
					fmt.Printf("itcMac = %d\n", itcMac)
					if itcMac > 0 && len(operand) >= 1+2*(itcMac+1) {
						fmt.Printf("rgdxaCenter:")
						for c := 0; c <= itcMac; c++ {
							val := int16(binary.LittleEndian.Uint16(operand[1+c*2:]))
							fmt.Printf(" %d", val)
						}
						fmt.Println()
						fmt.Printf("Column widths:")
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
		done:
	}
}
