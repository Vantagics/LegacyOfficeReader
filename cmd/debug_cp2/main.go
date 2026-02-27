package main

import (
	"encoding/binary"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/helpers"
)

func main() {
	adaptor, err := cfb.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer adaptor.CloseFile()

	var root, wordDoc, table1 *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "WordDocument":
			wordDoc = dir
		case "1Table":
			table1 = dir
		}
	}

	wReader, _ := adaptor.OpenObject(wordDoc, root)
	wSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wData := make([]byte, wSize)
	wReader.Read(wData)

	tReader, _ := adaptor.OpenObject(table1, root)
	tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tData := make([]byte, tSize)
	tReader.Read(tData)

	// Parse FIB
	offset := 0x20
	csw := binary.LittleEndian.Uint16(wData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wData[offset:])
	fibRgLwStart := offset + 2
	ccpText := binary.LittleEndian.Uint32(wData[fibRgLwStart+3*4:])
	fmt.Printf("ccpText=%d\n", ccpText)
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wData[offset:])
	offset += 2

	readFcLcb := func(index int) uint32 {
		if int(cbRgFcLcb) <= index {
			return 0
		}
		off := offset + index*4
		return binary.LittleEndian.Uint32(wData[off:])
	}

	fcClx := readFcLcb(66)
	lcbClx := readFcLcb(67)

	// Parse piece table directly
	clxData := tData[fcClx : fcClx+lcbClx]
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
		break
	}
	plcPcdLen := binary.LittleEndian.Uint32(clxData[pos:])
	pos += 4
	plcPcdData := clxData[pos : pos+plcPcdLen]

	n := (plcPcdLen - 4) / 12
	fmt.Printf("Pieces: %d\n", n)

	// Read CPs and piece descriptors
	cps := make([]uint32, n+1)
	for i := uint32(0); i <= n; i++ {
		cps[i] = binary.LittleEndian.Uint32(plcPcdData[i*4:])
	}

	// Extract text piece by piece and track CPs
	var fullText strings.Builder
	for i := uint32(0); i < n; i++ {
		pdStart := (n+1)*4 + i*8
		fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])
		charCount := cps[i+1] - cps[i]

		var actualOffset uint32
		var isUnicode bool
		if fc&0x40000000 != 0 {
			actualOffset = (fc & ^uint32(0x40000000)) >> 1
			isUnicode = false
		} else {
			actualOffset = fc
			isUnicode = true
		}

		var byteCount uint32
		if isUnicode {
			byteCount = charCount * 2
		} else {
			byteCount = charCount
		}

		end := actualOffset + byteCount
		if uint64(end) > uint64(len(wData)) {
			break
		}
		fragment := wData[actualOffset:end]
		var text string
		if isUnicode {
			text = helpers.DecodeUTF16LE(fragment)
		} else {
			text = helpers.DecodeWithCodepage(fragment, 936)
		}

		fmt.Printf("Piece %d: CP[%d-%d] fc=%d unicode=%v chars=%d\n",
			i, cps[i], cps[i+1], actualOffset, isUnicode, charCount)

		// Show first few chars
		runes := []rune(text)
		for j := 0; j < len(runes) && j < 5; j++ {
			r := runes[j]
			cp := cps[i] + uint32(j)
			if r < 0x20 {
				fmt.Printf("  CP %d: \\x%02X\n", cp, r)
			} else {
				fmt.Printf("  CP %d: %c (U+%04X)\n", cp, r, r)
			}
		}
		if len(runes) > 5 {
			fmt.Printf("  ... (%d more)\n", len(runes)-5)
		}

		fullText.WriteString(text)
	}

	// Now check CPs 6, 9, 12, 15, 108 in the full text
	allRunes := []rune(fullText.String())
	fmt.Printf("\nFull text length: %d runes\n", len(allRunes))
	checkCPs := []int{6, 9, 12, 15, 108}
	for _, cp := range checkCPs {
		if cp < len(allRunes) {
			r := allRunes[cp]
			if r < 0x20 {
				fmt.Printf("CP %d: \\x%02X\n", cp, r)
			} else {
				fmt.Printf("CP %d: %c (U+%04X)\n", cp, r, r)
			}
		}
	}
}
