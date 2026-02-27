package main

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"
	"unicode/utf16"

	"github.com/shakinm/xlsReader/cfb"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	c, err := cfb.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "CFB: %v\n", err)
		return
	}

	dirs := c.GetDirs()
	var root, wordDocDir, tableDir *cfb.Directory
	for _, d := range dirs {
		name := d.Name()
		if name == "Root Entry" { root = d }
		if name == "WordDocument" { wordDocDir = d }
		if name == "1Table" || name == "0Table" {
			if tableDir == nil { tableDir = d }
		}
	}

	wdReader, _ := c.OpenObject(wordDocDir, root)
	wordDocData, _ := io.ReadAll(wdReader)
	tReader, _ := c.OpenObject(tableDir, root)
	tableData, _ := io.ReadAll(tReader)

	fcStshf := binary.LittleEndian.Uint32(wordDocData[0xA2:0xA6])
	lcbStshf := binary.LittleEndian.Uint32(wordDocData[0xA6:0xAA])
	stshData := tableData[fcStshf : fcStshf+lcbStshf]

	cbStshi := binary.LittleEndian.Uint16(stshData[0:2])
	cstd := binary.LittleEndian.Uint16(stshData[2:4])
	cbSTDBase := binary.LittleEndian.Uint16(stshData[4:6])

	pos := int(cbStshi) + 2
	for i := 0; i < int(cstd); i++ {
		if pos+2 > len(stshData) { break }
		cbStd := binary.LittleEndian.Uint16(stshData[pos:])
		pos += 2
		if cbStd == 0 { continue }
		stdData := stshData[pos : pos+int(cbStd)]
		pos += int(cbStd)

		if i != 0 && i != 44 && i != 48 {
			continue
		}

		word0 := binary.LittleEndian.Uint16(stdData[0:2])
		sti := word0 & 0x0FFF
		word1 := binary.LittleEndian.Uint16(stdData[2:4])
		stk := word1 & 0x000F
		istdBase := (word1 >> 4) & 0x0FFF

		// Get name
		nameOff := int(cbSTDBase)
		name := ""
		nameLen := uint16(0)
		if nameOff+2 <= len(stdData) {
			nameLen = binary.LittleEndian.Uint16(stdData[nameOff:])
			if nameLen > 0 && nameOff+2+int(nameLen)*2 <= len(stdData) {
				u16 := make([]uint16, nameLen)
				for j := 0; j < int(nameLen); j++ {
					u16[j] = binary.LittleEndian.Uint16(stdData[nameOff+2+j*2:])
				}
				name = string(utf16.Decode(u16))
			}
		}

		fmt.Printf("\nStyle[%d] %q sti=%d stk=%d base=%d cbStd=%d\n", i, name, sti, stk, istdBase, cbStd)

		// Find UPX data
		upxOffset := nameOff + 2 + int(nameLen)*2 + 2
		if upxOffset%2 != 0 { upxOffset++ }

		if stk == 1 && upxOffset+2 <= len(stdData) {
			// First UPX: paragraph properties
			cbUpx := int(binary.LittleEndian.Uint16(stdData[upxOffset:]))
			upxOffset += 2
			fmt.Printf("  Para UPX: cbUpx=%d\n", cbUpx)
			if cbUpx > 0 && upxOffset+cbUpx <= len(stdData) {
				upxData := stdData[upxOffset : upxOffset+cbUpx]
				fmt.Printf("  Para UPX raw: ")
				for _, b := range upxData {
					fmt.Printf("%02X ", b)
				}
				fmt.Println()

				// Parse sprms (skip first 2 bytes = istd)
				if cbUpx > 2 {
					sprmData := upxData[2:]
					spos := 0
					for spos+2 <= len(sprmData) {
						op := binary.LittleEndian.Uint16(sprmData[spos:])
						spos += 2
						sprmType := (op >> 13) & 0x7
						var opSize int
						switch sprmType {
						case 0, 1: opSize = 1
						case 2, 4, 5: opSize = 2
						case 3: opSize = 4
						case 6:
							if spos < len(sprmData) {
								opSize = int(sprmData[spos]) + 1
							}
						case 7: opSize = 3
						}
						if spos+opSize > len(sprmData) { break }
						operand := sprmData[spos : spos+opSize]
						fmt.Printf("    sprm 0x%04X operand=", op)
						for _, b := range operand {
							fmt.Printf("%02X ", b)
						}
						if op == 0x2461 || op == 0x2403 {
							fmt.Printf(" (alignment=%d)", operand[0])
						}
						fmt.Println()
						spos += opSize
					}
				}
			}
			if cbUpx > 0 { upxOffset += cbUpx }
			if upxOffset%2 != 0 { upxOffset++ }

			// Second UPX: character properties
			if upxOffset+2 <= len(stdData) {
				cbUpx2 := int(binary.LittleEndian.Uint16(stdData[upxOffset:]))
				upxOffset += 2
				fmt.Printf("  Char UPX: cbUpx=%d\n", cbUpx2)
				if cbUpx2 > 0 && upxOffset+cbUpx2 <= len(stdData) {
					upxData := stdData[upxOffset : upxOffset+cbUpx2]
					fmt.Printf("  Char UPX raw: ")
					for _, b := range upxData {
						fmt.Printf("%02X ", b)
					}
					fmt.Println()
				}
			}
		}
	}
}
