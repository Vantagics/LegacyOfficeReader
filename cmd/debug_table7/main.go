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

	// Look at page 31 specifically (from earlier debug output)
	targetPages := []uint32{31, 34, 36}
	for i := uint32(0); i < n; i++ {
		pnOffset := (n+1)*4 + i*4
		pn := binary.LittleEndian.Uint32(plcData[pnOffset:])
		
		isTarget := false
		for _, tp := range targetPages {
			if pn == tp {
				isTarget = true
				break
			}
		}
		if !isTarget {
			continue
		}

		pageOffset := pn * 512
		if pageOffset+512 > uint32(len(wordDocData)) {
			continue
		}
		page := wordDocData[pageOffset : pageOffset+512]
		crun := int(page[511])
		fmt.Printf("=== Page %d, crun=%d ===\n", pn, crun)

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

			// Check if this has table-related sprms
			hasTable := false
			for k := 0; k+1 < len(sprmData); k++ {
				op := binary.LittleEndian.Uint16(sprmData[k:])
				if op == 0x2416 || op == 0x2417 || op == 0xD608 || op == 0xD605 {
					hasTable = true
					break
				}
			}
			if !hasTable {
				continue
			}

			fmt.Printf("\nRun %d: istd=%d, cb=%d, pos=%d, sprmData len=%d\n", j, istd, cb, pos, len(sprmData))
			
			// Walk through sprms step by step
			spos := 0
			for spos+2 <= len(sprmData) {
				opcode := binary.LittleEndian.Uint16(sprmData[spos:])
				opcodePos := spos
				spos += 2
				
				spra := (opcode >> 13) & 0x07
				var opSize int
				var sizeBytes int
				
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
						fmt.Printf("  @%d: opcode=0x%04X spra=%d - no size byte\n", opcodePos, opcode, spra)
						goto done
					}
					opSize = int(sprmData[spos])
					sizeBytes = 1
					spos++
				case 7:
					opSize = 3
				}
				
				if spos+opSize > len(sprmData) {
					fmt.Printf("  @%d: opcode=0x%04X spra=%d opSize=%d - truncated\n", opcodePos, opcode, spra, opSize)
					break
				}
				
				name := ""
				switch opcode {
				case 0x2416:
					name = "sprmPFInTable"
				case 0x2417:
					name = "sprmPFTtp"
				case 0xD608:
					name = "sprmTDefTable"
				case 0xD605:
					name = "sprmTDefTable10"
				case 0x2461:
					name = "sprmPJc"
				case 0x4600:
					name = "sprmPIstd"
				case 0x646B:
					name = "sprmTTableBorders80"
				}
				
				if name != "" || opcode == 0xD608 || opcode == 0xD605 {
					operand := sprmData[spos : spos+opSize]
					fmt.Printf("  @%d: opcode=0x%04X (%s) spra=%d sizeBytes=%d opSize=%d", opcodePos, opcode, name, spra, sizeBytes, opSize)
					if opSize <= 20 {
						fmt.Printf(" data=")
						for _, b := range operand {
							fmt.Printf("%02x ", b)
						}
					}
					fmt.Println()
					
					if opcode == 0xD608 && opSize > 0 {
						itcMac := int(operand[0])
						fmt.Printf("    itcMac=%d\n", itcMac)
						// Also try interpreting with 2-byte length
						fmt.Printf("    If 2-byte len: first 2 bytes = 0x%02x%02x = %d\n", operand[1], operand[0], binary.LittleEndian.Uint16(operand[0:2]))
					}
				}
				
				spos += opSize
			}
			done:
		}
	}
}
