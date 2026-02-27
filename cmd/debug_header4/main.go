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
	fcPlcfHdd := readFcLcb(22)
	lcbPlcfHdd := readFcLcb(23)

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

	type piece struct {
		cpStart, cpEnd uint32
		fc             uint32
		isUnicode      bool
	}
	cps := make([]uint32, n+1)
	for i := uint32(0); i <= n; i++ {
		cps[i] = binary.LittleEndian.Uint32(plcPcdData[i*4:])
	}
	cpArraySize := (n + 1) * 4
	pieces := make([]piece, n)
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
		pieces[i] = piece{cpStart: cps[i], cpEnd: cps[i+1], fc: actualOffset, isUnicode: isUnicode}
	}

	hddStart := ccpText + ccpFtn
	hddEnd := hddStart + ccpHdd

	fmt.Printf("ccpText=%d ccpFtn=%d ccpHdd=%d\n", ccpText, ccpFtn, ccpHdd)
	fmt.Printf("Header text CP range: %d to %d\n", hddStart, hddEnd)

	// Extract the raw header text character by character
	fmt.Printf("\nRaw header/footer text (CP %d to %d):\n", hddStart, hddEnd)
	for cp := hddStart; cp < hddEnd; cp++ {
		// Find the piece containing this CP
		for _, p := range pieces {
			if cp >= p.cpStart && cp < p.cpEnd {
				charOffset := cp - p.cpStart
				var byteOffset uint32
				if p.isUnicode {
					byteOffset = p.fc + charOffset*2
				} else {
					byteOffset = p.fc + charOffset
				}
				if p.isUnicode && uint64(byteOffset)+2 <= uint64(len(wordDocData)) {
					ch := helpers.DecodeUTF16LE(wordDocData[byteOffset : byteOffset+2])
					r := []rune(ch)
					if len(r) > 0 {
						if r[0] < 0x20 && r[0] != '\t' {
							fmt.Printf("CP%d: [%02X]\n", cp, r[0])
						} else {
							fmt.Printf("CP%d: U+%04X %c\n", cp, r[0], r[0])
						}
					}
				} else if !p.isUnicode && uint64(byteOffset)+1 <= uint64(len(wordDocData)) {
					b := wordDocData[byteOffset]
					decoded := helpers.DecodeWithCodepage([]byte{b}, 936)
					r := []rune(decoded)
					if len(r) > 0 {
						fmt.Printf("CP%d: 0x%02X -> %c\n", cp, b, r[0])
					}
				}
				break
			}
		}
	}

	// Show PlcfHdd entries with story text
	fmt.Printf("\nPlcfHdd stories:\n")
	plcHddData := tableData[fcPlcfHdd : fcPlcfHdd+lcbPlcfHdd]
	nCPs := lcbPlcfHdd / 4
	storyNames := []string{"even header", "odd header", "even footer", "odd footer", "first header", "first footer"}
	for i := uint32(0); i+1 < nCPs; i++ {
		cpS := binary.LittleEndian.Uint32(plcHddData[i*4:])
		cpE := binary.LittleEndian.Uint32(plcHddData[(i+1)*4:])
		storyName := "unknown"
		idx := i % 6
		if int(idx) < len(storyNames) {
			storyName = storyNames[idx]
		}
		fmt.Printf("  Story %d (%s): CP[%d-%d] (abs %d-%d) len=%d\n",
			i, storyName, cpS, cpE, hddStart+cpS, hddStart+cpE, cpE-cpS)
	}

	_ = fcPlcfHdd
	_ = lcbPlcfHdd
}
