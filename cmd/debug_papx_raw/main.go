package main

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"

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
		os.Exit(1)
	}

	dirs := c.GetDirs()
	var root *cfb.Directory
	var wordDocDir *cfb.Directory
	for _, d := range dirs {
		name := d.Name()
		if name == "Root Entry" {
			root = d
		}
		if name == "WordDocument" {
			wordDocDir = d
		}
	}

	wdReader, err := c.OpenObject(wordDocDir, root)
	if err != nil {
		fmt.Fprintf(os.Stderr, "WD: %v\n", err)
		return
	}
	wordDocData, _ := io.ReadAll(wdReader)

	// Read FIB to get PlcBtePapx
	fcPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x102:0x106])
	lcbPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x106:0x10A])

	// Also need table stream
	f.Seek(0, 0)
	c2, _ := cfb.OpenReader(f)
	dirs2 := c2.GetDirs()
	var root2, tableDir *cfb.Directory
	for _, d := range dirs2 {
		name := d.Name()
		if name == "Root Entry" {
			root2 = d
		}
		if name == "1Table" || name == "0Table" {
			if tableDir == nil {
				tableDir = d
			}
		}
	}
	tReader, _ := c2.OpenObject(tableDir, root2)
	tableData, _ := io.ReadAll(tReader)

	fmt.Printf("PlcBtePapx: fc=%d lcb=%d\n", fcPlcfBtePapx, lcbPlcfBtePapx)

	plcData := tableData[fcPlcfBtePapx : fcPlcfBtePapx+lcbPlcfBtePapx]
	n := (lcbPlcfBtePapx - 4) / 8

	// Look at a few PAPX pages to find alignment sprms
	alignSprms := map[uint16]int{}
	for i := uint32(0); i < n && i < 20; i++ {
		pnOffset := (n+1)*4 + i*4
		pn := binary.LittleEndian.Uint32(plcData[pnOffset:])
		pageOffset := pn * 512
		if pageOffset+512 > uint32(len(wordDocData)) {
			continue
		}
		page := wordDocData[pageOffset : pageOffset+512]
		crun := int(page[511])

		fcArraySize := (crun + 1) * 4
		bxStart := fcArraySize

		for j := 0; j < crun; j++ {
			bxPos := bxStart + j*13
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

			// Scan for alignment sprms
			spos := 0
			for spos+2 <= len(sprmData) {
				op := binary.LittleEndian.Uint16(sprmData[spos:])
				spos += 2

				if op == 0x2461 || op == 0x2403 {
					if spos < len(sprmData) {
						val := sprmData[spos]
						alignSprms[op]++
						if i < 5 && j < 3 {
							fmt.Printf("  Page %d Run %d: sprm 0x%04X align=%d\n", i, j, op, val)
						}
						spos++
					}
				} else {
					// Skip operand
					sprmType := (op >> 13) & 0x7
					switch sprmType {
					case 0, 1:
						spos++
					case 2, 4, 5:
						spos += 2
					case 3:
						spos += 4
					case 6:
						if op == 0xD608 {
							if spos+2 <= len(sprmData) {
								sz := int(binary.LittleEndian.Uint16(sprmData[spos:]))
								spos += 2 + sz
							} else {
								spos = len(sprmData)
							}
						} else {
							if spos < len(sprmData) {
								spos += int(sprmData[spos]) + 1
							}
						}
					case 7:
						spos += 3
					}
				}
			}
		}
	}

	fmt.Printf("\nAlignment sprm counts:\n")
	for op, count := range alignSprms {
		fmt.Printf("  0x%04X: %d occurrences\n", op, count)
	}
}
