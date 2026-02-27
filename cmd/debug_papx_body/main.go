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

	fcPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x102:0x106])
	lcbPlcfBtePapx := binary.LittleEndian.Uint32(wordDocData[0x106:0x10A])
	plcData := tableData[fcPlcfBtePapx : fcPlcfBtePapx+lcbPlcfBtePapx]
	n := (lcbPlcfBtePapx - 4) / 8

	// Scan ALL PAPX runs and count alignment values
	alignCounts := map[string]int{}
	totalRuns := 0
	noAlignRuns := 0

	for i := uint32(0); i < n; i++ {
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
			totalRuns++
			bxPos := bxStart + j*13
			if bxPos >= 511 { break }
			bOffset := int(page[bxPos])
			if bOffset == 0 {
				noAlignRuns++
				continue
			}

			pos := bOffset * 2
			if pos >= 512 { continue }

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
					if endPos > 512 { endPos = 512 }
					if pos+4 < endPos {
						sprmData = page[pos+4 : endPos]
					}
				}
			} else {
				dataLen := cb*2 - 1
				if dataLen >= 2 && pos+3 <= 512 {
					istd = binary.LittleEndian.Uint16(page[pos+1 : pos+3])
				}
				endPos := pos + 1 + dataLen
				if endPos > 512 { endPos = 512 }
				if pos+3 < endPos {
					sprmData = page[pos+3 : endPos]
				}
			}

			// Check for alignment sprms
			hasAlign := false
			spos := 0
			for spos+2 <= len(sprmData) {
				op := binary.LittleEndian.Uint16(sprmData[spos:])
				spos += 2

				if op == 0x2461 {
					if spos < len(sprmData) {
						val := sprmData[spos]
						key := fmt.Sprintf("istd=%d align=%d", istd, val)
						alignCounts[key]++
						hasAlign = true
						spos++
					}
				} else {
					sprmType := (op >> 13) & 0x7
					switch sprmType {
					case 0, 1: spos++
					case 2, 4, 5: spos += 2
					case 3: spos += 4
					case 6:
						if op == 0xD608 {
							if spos+2 <= len(sprmData) {
								sz := int(binary.LittleEndian.Uint16(sprmData[spos:]))
								spos += 2 + sz
							} else { spos = len(sprmData) }
						} else {
							if spos < len(sprmData) {
								spos += int(sprmData[spos]) + 1
							}
						}
					case 7: spos += 3
					}
				}
			}
			if !hasAlign {
				noAlignRuns++
			}
		}
	}

	fmt.Printf("Total PAPX runs: %d\n", totalRuns)
	fmt.Printf("Runs without alignment sprm: %d\n", noAlignRuns)
	fmt.Printf("\nAlignment by istd:\n")
	for key, count := range alignCounts {
		fmt.Printf("  %s: %d\n", key, count)
	}
}
