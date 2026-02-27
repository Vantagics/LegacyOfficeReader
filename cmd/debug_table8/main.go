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

	// Look at page 31 run 0 - dump the full sprm data around sprmTDefTable
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
		
		// cb=238, so dataLen = 238*2-1 = 475
		dataLen := cb*2 - 1
		istd := binary.LittleEndian.Uint16(page[pos+1 : pos+3])
		endPos := pos + 1 + dataLen
		if endPos > 512 {
			endPos = 512
		}
		sprmData := page[pos+3 : endPos]
		
		fmt.Printf("Page 31, run 0: cb=%d, dataLen=%d, istd=%d, sprmData len=%d\n", cb, dataLen, istd, len(sprmData))
		fmt.Printf("pos=%d, endPos=%d\n\n", pos, endPos)
		
		// Walk sprms to find sprmTDefTable
		spos := 0
		for spos+2 <= len(sprmData) {
			opcode := binary.LittleEndian.Uint16(sprmData[spos:])
			opcodePos := spos
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
				if spos >= len(sprmData) {
					break
				}
				opSize = int(sprmData[spos])
				spos++
			case 7:
				opSize = 3
			}
			
			if spos+opSize > len(sprmData) {
				break
			}
			
			if opcode == 0xD608 {
				operand := sprmData[spos : spos+opSize]
				fmt.Printf("sprmTDefTable at offset %d, opSize=%d\n", opcodePos, opSize)
				fmt.Printf("Operand hex dump:\n")
				for k := 0; k < len(operand); k++ {
					fmt.Printf("%02x ", operand[k])
					if (k+1)%16 == 0 {
						fmt.Println()
					}
				}
				fmt.Println()
				
				// Try interpreting: first byte is itcMac
				fmt.Printf("\nitcMac (byte 0) = %d\n", operand[0])
				
				// What if the first 2 bytes are a length prefix?
				// Then the actual TDefTable data starts at byte 2
				if len(operand) >= 3 {
					altItcMac := int(operand[2])
					fmt.Printf("Alternative: skip 2 bytes, itcMac = %d\n", altItcMac)
					if altItcMac > 0 && len(operand) >= 3+2*(altItcMac+1) {
						fmt.Printf("  rgdxaCenter:")
						for c := 0; c <= altItcMac; c++ {
							val := int16(binary.LittleEndian.Uint16(operand[3+c*2:]))
							fmt.Printf(" %d", val)
						}
						fmt.Println()
						fmt.Printf("  Column widths:")
						for c := 0; c < altItcMac; c++ {
							left := int16(binary.LittleEndian.Uint16(operand[3+c*2:]))
							right := int16(binary.LittleEndian.Uint16(operand[3+(c+1)*2:]))
							fmt.Printf(" %d", right-left)
						}
						fmt.Println()
					}
				}
			}
			
			if opcode == 0xD605 {
				operand := sprmData[spos : spos+opSize]
				fmt.Printf("sprmTDefTable10 at offset %d, opSize=%d\n", opcodePos, opSize)
				fmt.Printf("Operand hex dump:\n")
				for k := 0; k < len(operand); k++ {
					fmt.Printf("%02x ", operand[k])
					if (k+1)%16 == 0 {
						fmt.Println()
					}
				}
				fmt.Println()
				fmt.Printf("itcMac (byte 0) = %d\n", operand[0])
				if operand[0] > 0 {
					itcMac := int(operand[0])
					if len(operand) >= 1+2*(itcMac+1) {
						fmt.Printf("  rgdxaCenter:")
						for c := 0; c <= itcMac; c++ {
							val := int16(binary.LittleEndian.Uint16(operand[1+c*2:]))
							fmt.Printf(" %d", val)
						}
						fmt.Println()
					}
				}
				fmt.Println()
			}
			
			spos += opSize
		}
	}
}
