package main

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/helpers"
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

	// Parse piece table manually
	fcClx := binary.LittleEndian.Uint32(wordDocData[0x1A2:0x1A6])
	lcbClx := binary.LittleEndian.Uint32(wordDocData[0x1A6:0x1AA])
	clxData := tableData[fcClx : fcClx+lcbClx]

	pos := uint32(0)
	for pos < uint32(len(clxData)) {
		if clxData[pos] == 0x01 {
			prcSize := binary.LittleEndian.Uint16(clxData[pos+1:])
			pos += 3 + uint32(prcSize)
			continue
		}
		if clxData[pos] == 0x02 {
			pos++
			break
		}
		pos++
	}

	plcPcdLen := binary.LittleEndian.Uint32(clxData[pos:])
	pos += 4
	plcPcdData := clxData[pos : pos+plcPcdLen]
	n := (plcPcdLen - 4) / 12

	// Extract text from all pieces
	totalChars := 0
	for i := uint32(0); i < n; i++ {
		cpStart := binary.LittleEndian.Uint32(plcPcdData[i*4:])
		cpEnd := binary.LittleEndian.Uint32(plcPcdData[(i+1)*4:])
		pdStart := (n+1)*4 + i*8
		fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])
		isUnicode := fc&0x40000000 == 0

		charCount := cpEnd - cpStart
		var actualFC uint32
		if isUnicode {
			actualFC = fc
		} else {
			actualFC = (fc & ^uint32(0x40000000)) >> 1
		}

		var byteCount uint32
		if isUnicode {
			byteCount = charCount * 2
		} else {
			byteCount = charCount
		}

		end := actualFC + byteCount
		if uint64(end) > uint64(len(wordDocData)) {
			fmt.Printf("Piece[%d]: cp[%d-%d] OUT OF BOUNDS (fc=%d, end=%d, streamLen=%d)\n",
				i, cpStart, cpEnd, actualFC, end, len(wordDocData))
			continue
		}

		fragment := wordDocData[actualFC:end]
		var text string
		if isUnicode {
			text = helpers.DecodeUTF16LE(fragment)
		} else {
			text = helpers.DecodeANSI(fragment)
		}

		runes := []rune(text)
		totalChars += len(runes)
		fmt.Printf("Piece[%d]: cp[%d-%d] chars=%d unicode=%v\n", i, cpStart, cpEnd, len(runes), isUnicode)
	}

	fmt.Printf("\nTotal chars extracted: %d\n", totalChars)

	// Now show the header area text
	ccpText := binary.LittleEndian.Uint32(wordDocData[0x4C:0x50])
	ccpHdd := binary.LittleEndian.Uint32(wordDocData[0x54:0x58])
	fmt.Printf("ccpText=%d ccpHdd=%d\n", ccpText, ccpHdd)
	fmt.Printf("Header area: CP %d to %d\n", ccpText, ccpText+ccpHdd)

	// Extract full text and show header area
	fullText := ""
	for i := uint32(0); i < n; i++ {
		cpStart := binary.LittleEndian.Uint32(plcPcdData[i*4:])
		cpEnd := binary.LittleEndian.Uint32(plcPcdData[(i+1)*4:])
		pdStart := (n+1)*4 + i*8
		fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])
		isUnicode := fc&0x40000000 == 0

		charCount := cpEnd - cpStart
		var actualFC uint32
		if isUnicode {
			actualFC = fc
		} else {
			actualFC = (fc & ^uint32(0x40000000)) >> 1
		}

		var byteCount uint32
		if isUnicode {
			byteCount = charCount * 2
		} else {
			byteCount = charCount
		}

		end := actualFC + byteCount
		if uint64(end) > uint64(len(wordDocData)) {
			continue
		}

		fragment := wordDocData[actualFC:end]
		if isUnicode {
			fullText += helpers.DecodeUTF16LE(fragment)
		} else {
			fullText += helpers.DecodeANSI(fragment)
		}
		_ = cpStart
	}

	runes := []rune(fullText)
	fmt.Printf("Full text runes: %d\n", len(runes))

	// Show header area
	hdrStart := int(ccpText)
	hdrEnd := int(ccpText + ccpHdd)
	if hdrEnd > len(runes) {
		hdrEnd = len(runes)
	}
	if hdrStart < len(runes) {
		fmt.Printf("\nHeader text (CP %d-%d):\n", hdrStart, hdrEnd)
		for i := hdrStart; i < hdrEnd; i++ {
			r := runes[i]
			if r < 0x20 {
				fmt.Printf("[%02X]", r)
			} else {
				fmt.Printf("%c", r)
			}
		}
		fmt.Println()
	} else {
		fmt.Printf("Header area starts beyond text (hdrStart=%d, textLen=%d)\n", hdrStart, len(runes))
	}
}
