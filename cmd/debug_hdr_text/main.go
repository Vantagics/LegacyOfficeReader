package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/doc"
)

func main() {
	adaptor, err := cfb.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer adaptor.CloseFile()

	var wordDoc, table0, table1, root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "WordDocument":
			wordDoc = dir
		case "0Table":
			table0 = dir
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

	flags := binary.LittleEndian.Uint16(wordDocData[0x0A:])
	fWhichTblStm := (flags >> 9) & 1
	var tableDir *cfb.Directory
	if fWhichTblStm == 1 {
		tableDir = table1
	} else {
		tableDir = table0
	}

	tableReader, _ := adaptor.OpenObject(tableDir, root)
	tableSize := binary.LittleEndian.Uint32(tableDir.StreamSize[:])
	tableData := make([]byte, tableSize)
	tableReader.Read(tableData)

	// Use the doc package to get the full text
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fullText := d.GetText()
	runes := []rune(fullText)
	fmt.Printf("Main text length (runes): %d\n", len(runes))

	// FIB fields
	offset := 0x20
	csw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2
	fibRgLwStart := offset

	ccpText := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+3*4:])
	ccpFtn := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+4*4:])
	ccpHdd := binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+5*4:])

	fmt.Printf("ccpText=%d ccpFtn=%d ccpHdd=%d\n", ccpText, ccpFtn, ccpHdd)

	offset += int(cslw) * 4
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

	fcPlcfHdd := readFcLcb(22)
	lcbPlcfHdd := readFcLcb(23)
	fmt.Printf("fcPlcfHdd=%d lcbPlcfHdd=%d\n", fcPlcfHdd, lcbPlcfHdd)

	// The full text from GetText() is only the main body text (ccpText chars).
	// We need the FULL character stream including header/footer text.
	// The header text starts at ccpText + ccpFtn in the full stream.
	// But GetText() only returns ccpText chars.
	// So we need to look at the raw text extraction.
	
	// Let's check: the full text should be ccpText + ccpFtn + ccpHdd + 1 chars
	expectedTotal := ccpText + ccpFtn + ccpHdd + 1
	fmt.Printf("Expected total chars: %d, got main text: %d\n", expectedTotal, len(runes))
	
	// The header/footer text is at positions ccpText+ccpFtn to ccpText+ccpFtn+ccpHdd
	hddStart := ccpText + ccpFtn
	hddEnd := hddStart + ccpHdd
	fmt.Printf("Header/footer text range: [%d, %d)\n", hddStart, hddEnd)
	
	// Since GetText() only returns main body text, the header text is NOT in fullText.
	// The extractHeaderFooter function in header.go receives the FULL text.
	// But doc.go passes fullText which is built from pieces.
	// Let's check if the piece table covers the header area.
	
	fmt.Printf("\nMain text ends at CP %d\n", ccpText)
	fmt.Printf("Header text should be at CP [%d, %d)\n", hddStart, hddEnd)

	// Parse PlcfHdd from table stream
	if lcbPlcfHdd > 0 && uint64(fcPlcfHdd)+uint64(lcbPlcfHdd) <= uint64(len(tableData)) {
		plcData := tableData[fcPlcfHdd : fcPlcfHdd+lcbPlcfHdd]
		nCPs := lcbPlcfHdd / 4
		fmt.Printf("\nPlcfHdd: %d CPs (%d stories)\n", nCPs, nCPs-1)

		cps := make([]uint32, nCPs)
		for i := uint32(0); i < nCPs; i++ {
			cps[i] = binary.LittleEndian.Uint32(plcData[i*4:])
		}

		// Print first 20 CPs
		for i := 0; i < int(nCPs) && i < 20; i++ {
			storyIdx := i % 6
			storyName := ""
			switch storyIdx {
			case 0:
				storyName = "even-hdr"
			case 1:
				storyName = "odd-hdr"
			case 2:
				storyName = "even-ftr"
			case 3:
				storyName = "odd-ftr"
			case 4:
				storyName = "first-hdr"
			case 5:
				storyName = "first-ftr"
			}
			cpStart := cps[i]
			cpEnd := uint32(0)
			if i+1 < int(nCPs) {
				cpEnd = cps[i+1]
			}
			absStart := hddStart + cpStart
			absEnd := hddStart + cpEnd
			fmt.Printf("  Story[%d] (%s): relCP[%d-%d] absCP[%d-%d] len=%d\n",
				i, storyName, cpStart, cpEnd, absStart, absEnd, cpEnd-cpStart)
			
			// If within range, show the text
			if absStart < uint32(len(runes)) && absEnd <= uint32(len(runes)) && cpEnd > cpStart && cpEnd-cpStart < 200 {
				text := string(runes[absStart:absEnd])
				fmt.Printf("    text: %q\n", text)
			}
		}
	}
}
