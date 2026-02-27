package main

import (
	"encoding/binary"
	"fmt"
	"strings"

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

	// Parse piece table and extract full text
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

	var sb strings.Builder
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
			if isUnicode {
				sb.WriteString(helpers.DecodeUTF16LE(fragment))
			} else {
				sb.WriteString(helpers.DecodeWithCodepage(fragment, 936))
			}
		}
	}

	fullText := sb.String()
	runes := []rune(fullText)
	fmt.Printf("Full text rune count: %d\n", len(runes))
	fmt.Printf("ccpText=%d ccpFtn=%d ccpHdd=%d\n", ccpText, ccpFtn, ccpHdd)

	hddStart := ccpText + ccpFtn
	fmt.Printf("hddStart=%d\n", hddStart)

	// Now simulate extractHeaderFooter
	plcHddData := tableData[fcPlcfHdd : fcPlcfHdd+lcbPlcfHdd]
	nCPs2 := lcbPlcfHdd / 4
	hddCPs := make([]uint32, nCPs2)
	for i := uint32(0); i < nCPs2; i++ {
		hddCPs[i] = binary.LittleEndian.Uint32(plcHddData[i*4:])
	}

	storyNames := []string{"even hdr", "odd hdr", "even ftr", "odd ftr", "first hdr", "first ftr"}
	for i := 0; i+1 < int(nCPs2); i++ {
		cpStart := hddStart + hddCPs[i]
		cpEnd := hddStart + hddCPs[i+1]

		storyName := "unknown"
		idx := i % 6
		if idx < len(storyNames) {
			storyName = storyNames[idx]
		}

		if cpStart >= uint32(len(runes)) || cpEnd > uint32(len(runes)) || cpStart >= cpEnd {
			fmt.Printf("Story %d (%s): CP[%d-%d] SKIP (out of range)\n", i, storyName, cpStart, cpEnd)
			continue
		}

		storyText := string(runes[cpStart:cpEnd])
		trimmed := strings.TrimRight(storyText, "\r\n")

		// Show raw chars
		fmt.Printf("Story %d (%s): CP[%d-%d] raw=", i, storyName, cpStart, cpEnd)
		for _, r := range []rune(trimmed) {
			if r < 0x20 && r != '\t' {
				fmt.Printf("[%02X]", r)
			} else {
				fmt.Printf("%c", r)
			}
		}
		fmt.Println()
	}
}
