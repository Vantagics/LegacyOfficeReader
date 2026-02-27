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
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2

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

	// Parse piece table
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

	cps := make([]uint32, n+1)
	for i := uint32(0); i <= n; i++ {
		cps[i] = binary.LittleEndian.Uint32(plcPcdData[i*4:])
	}
	cpArraySize := (n + 1) * 4

	// Extract full text using all pieces
	var totalChars uint32
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
		totalChars += charCount

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
			fmt.Printf("Piece %d: CP[%d-%d] %d chars, decoded %d runes\n",
				i, cps[i], cps[i+1], charCount, len(runes))
		}
	}
	fmt.Printf("Total chars from pieces: %d (last CP: %d)\n", totalChars, cps[n])

	// Now extract text at CP 10327-10345 (header area)
	fmt.Printf("\nHeader area text (CP 10327-10345):\n")
	for cp := uint32(10327); cp < 10345; cp++ {
		for i := uint32(0); i < n; i++ {
			if cp >= cps[i] && cp < cps[i+1] {
				pdStart := cpArraySize + i*8
				fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])
				isUnicode := fc&0x40000000 == 0
				var actualOffset uint32
				if isUnicode {
					actualOffset = fc
				} else {
					actualOffset = (fc & ^uint32(0x40000000)) >> 1
				}
				charOffset := cp - cps[i]
				var byteOffset uint32
				if isUnicode {
					byteOffset = actualOffset + charOffset*2
				} else {
					byteOffset = actualOffset + charOffset
				}
				if isUnicode && uint64(byteOffset)+2 <= uint64(len(wordDocData)) {
					ch := helpers.DecodeUTF16LE(wordDocData[byteOffset : byteOffset+2])
					r := []rune(ch)
					if len(r) > 0 {
						if r[0] < 0x20 && r[0] != '\t' {
							fmt.Printf("  CP%d: [%02X] (piece %d)\n", cp, r[0], i)
						} else {
							fmt.Printf("  CP%d: U+%04X %c (piece %d)\n", cp, r[0], r[0], i)
						}
					}
				}
				break
			}
		}
	}
}
