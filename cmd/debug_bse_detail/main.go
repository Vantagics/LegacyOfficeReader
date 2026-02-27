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

	// Find DggContainer
	var dggStart, dggEnd uint32
	for i := uint32(0); i+8 <= uint32(len(tableData)); i++ {
		vi := binary.LittleEndian.Uint16(tableData[i : i+2])
		rv := vi & 0x0F
		rt := binary.LittleEndian.Uint16(tableData[i+2 : i+4])
		rl := binary.LittleEndian.Uint32(tableData[i+4 : i+8])
		if rt == 0xF000 && rv == 0xF && rl > 0 && i+8+rl <= uint32(len(tableData)) {
			dggStart = i + 8
			dggEnd = i + 8 + rl
			fmt.Printf("DggContainer at offset %d, len=%d\n", i, rl)
			break
		}
	}

	// Find BStoreContainer
	var bscStart, bscEnd uint32
	for offset := dggStart; offset+8 <= dggEnd; {
		rt := binary.LittleEndian.Uint16(tableData[offset+2 : offset+4])
		rl := binary.LittleEndian.Uint32(tableData[offset+4 : offset+8])
		vi := binary.LittleEndian.Uint16(tableData[offset : offset+2])
		rv := vi & 0x0F
		inst := vi >> 4

		childEnd := offset + 8 + rl
		if childEnd > dggEnd {
			break
		}

		if rt == 0xF001 && rv == 0xF {
			bscStart = offset + 8
			bscEnd = childEnd
			fmt.Printf("BStoreContainer at offset %d, len=%d, inst=%d (num BSE entries)\n", offset, rl, inst)
			break
		}
		offset = childEnd
	}

	// Parse BSE entries
	bseIdx := 0
	for offset := bscStart; offset+8 <= bscEnd; {
		vi := binary.LittleEndian.Uint16(tableData[offset : offset+2])
		rt := binary.LittleEndian.Uint16(tableData[offset+2 : offset+4])
		rl := binary.LittleEndian.Uint32(tableData[offset+4 : offset+8])
		inst := vi >> 4

		childEnd := offset + 8 + rl
		if childEnd > bscEnd {
			break
		}

		if rt == 0xF007 {
			// BSE record header is 36 bytes
			if offset+8+36 <= bscEnd {
				bseData := tableData[offset+8 : offset+8+36]
				btWin32 := bseData[0]
				btMacOS := bseData[1]
				// rgbUid is bytes 2-17 (16 bytes)
				tag := binary.LittleEndian.Uint16(bseData[18:20])
				size := binary.LittleEndian.Uint32(bseData[20:24])
				cRef := binary.LittleEndian.Uint32(bseData[24:28])
				foDelay := binary.LittleEndian.Uint32(bseData[28:32])

				btName := "unknown"
				switch btWin32 {
				case 2:
					btName = "EMF"
				case 3:
					btName = "WMF"
				case 4:
					btName = "PICT"
				case 5:
					btName = "JPEG"
				case 6:
					btName = "PNG"
				case 7:
					btName = "DIB"
				case 8:
					btName = "TIFF"
				}

				fmt.Printf("\nBSE[%d]: inst=%d, btWin32=%d(%s), btMacOS=%d, tag=%d, size=%d, cRef=%d, foDelay=%d\n",
					bseIdx, inst, btWin32, btName, btMacOS, tag, size, cRef, foDelay)

				// Check if there's an embedded blip after the 36-byte header
				if rl > 36 {
					blipOffset := offset + 8 + 36
					if blipOffset+8 <= bscEnd {
						bvi := binary.LittleEndian.Uint16(tableData[blipOffset:])
						brt := binary.LittleEndian.Uint16(tableData[blipOffset+2:])
						brl := binary.LittleEndian.Uint32(tableData[blipOffset+4:])
						fmt.Printf("  Embedded blip: recType=0x%04X, inst=0x%04X, len=%d\n", brt, bvi>>4, brl)
					}
				}

				// If foDelay > 0, check what's at that offset in WordDocument stream
				if foDelay > 0 && foDelay+8 <= uint32(len(wordDocData)) {
					bvi := binary.LittleEndian.Uint16(wordDocData[foDelay:])
					brt := binary.LittleEndian.Uint16(wordDocData[foDelay+2:])
					brl := binary.LittleEndian.Uint32(wordDocData[foDelay+4:])
					fmt.Printf("  Delayed blip at WordDoc[%d]: recType=0x%04X, inst=0x%04X, len=%d\n", foDelay, brt, bvi>>4, brl)
				}
			}
			bseIdx++
		}

		offset = childEnd
	}
}
