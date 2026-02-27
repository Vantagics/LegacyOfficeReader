package main

import (
	"encoding/binary"
	"fmt"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/helpers"
)

func main() {
	adaptor, _ := cfb.OpenFile("testfie/test.doc")
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

	// Parse FIB
	offset := 0x20
	csw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wordDocData[offset:])
	fibRgLwStart := offset + 2
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2

	ccpText := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+3*4:])
	ccpFtn := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+4*4:])
	ccpHdd := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+5*4:])
	ccpAtn := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+6*4:])
	ccpEdn := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+7*4:])
	ccpTxbx := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+8*4:])
	ccpHdrTxbx := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+9*4:])

	totalCP := ccpText + ccpFtn + ccpHdd + ccpAtn + ccpEdn + ccpTxbx + ccpHdrTxbx
	if ccpFtn != 0 || ccpHdd != 0 || ccpAtn != 0 || ccpEdn != 0 || ccpTxbx != 0 || ccpHdrTxbx != 0 {
		totalCP++ // +1 for the final separator
	}

	fmt.Printf("ccpText=%d ccpFtn=%d ccpHdd=%d ccpAtn=%d ccpEdn=%d ccpTxbx=%d ccpHdrTxbx=%d\n",
		ccpText, ccpFtn, ccpHdd, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx)
	fmt.Printf("Total CP (calculated): %d\n", totalCP)

	// Parse piece table to see total CP range
	readFcLcb := func(index int) uint32 {
		if int(cbRgFcLcb) <= index {
			return 0
		}
		off := offset + index*4
		if off+4 > len(wordDocData) {
			return 0
		}
		return binary.LittleEndian.Uint32(wordDocData[off:])
	}

	fcClx := readFcLcb(66)
	lcbClx := readFcLcb(67)
	fmt.Printf("fcClx=%d lcbClx=%d\n", fcClx, lcbClx)

	clxData := tableData[fcClx : fcClx+lcbClx]
	pos := uint32(0)
	for pos < uint32(len(clxData)) {
		typeByte := clxData[pos]
		if typeByte == 0x01 {
			prcSize := binary.LittleEndian.Uint16(clxData[pos+1:])
			pos += 3 + uint32(prcSize)
			continue
		}
		if typeByte == 0x02 {
			pos++
			break
		}
		break
	}

	plcPcdLen := binary.LittleEndian.Uint32(clxData[pos:])
	pos += 4
	plcPcdData := clxData[pos : pos+plcPcdLen]

	n := (plcPcdLen - 4) / 12
	fmt.Printf("Piece table: %d pieces\n", n)

	cps := make([]uint32, n+1)
	for i := uint32(0); i <= n; i++ {
		cps[i] = binary.LittleEndian.Uint32(plcPcdData[i*4:])
	}

	cpArraySize := (n + 1) * 4
	for i := uint32(0); i < n; i++ {
		pdStart := cpArraySize + i*8
		fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])
		isUnicode := fc&0x40000000 == 0
		var actualOffset uint32
		if isUnicode {
			actualOffset = fc
		} else {
			actualOffset = (fc & ^uint32(0x40000000)) >> 1
		}
		charCount := cps[i+1] - cps[i]
		enc := "ANSI"
		if isUnicode {
			enc = "UTF16"
		}
		fmt.Printf("  Piece %d: CP[%d-%d] (%d chars) fc=0x%X %s offset=0x%X\n",
			i, cps[i], cps[i+1], charCount, fc, enc, actualOffset)

		// Show first few chars of this piece
		var byteCount uint32
		if isUnicode {
			byteCount = charCount * 2
		} else {
			byteCount = charCount
		}
		start := actualOffset
		end := start + byteCount
		if uint64(end) <= uint64(len(wordDocData)) {
			fragment := wordDocData[start:end]
			var text string
			if isUnicode {
				text = helpers.DecodeUTF16LE(fragment)
			} else {
				text = helpers.DecodeWithCodepage(fragment, 936)
			}
			runes := []rune(text)
			if len(runes) > 40 {
				runes = runes[:40]
			}
			var display string
			for _, r := range runes {
				if r < 0x20 && r != '\t' {
					display += fmt.Sprintf("[%02X]", r)
				} else {
					display += string(r)
				}
			}
			fmt.Printf("    text: %s\n", display)
		}
	}

	fmt.Printf("\nPiece table CP range: %d to %d\n", cps[0], cps[n])
}
