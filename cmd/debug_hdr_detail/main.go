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

	if wordDoc == nil {
		fmt.Println("No WordDocument stream")
		return
	}

	wordDocReader, _ := adaptor.OpenObject(wordDoc, root)
	wordDocSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wordDocData := make([]byte, wordDocSize)
	wordDocReader.Read(wordDocData)

	// Parse FIB basics
	// fWhichTblStm at bit 9 of flags at offset 0x000A
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

	// FIB fields
	ccpText := binary.LittleEndian.Uint32(wordDocData[0x004C:])
	ccpFtn := binary.LittleEndian.Uint32(wordDocData[0x0050:])
	ccpHdd := binary.LittleEndian.Uint32(wordDocData[0x0054:])

	fmt.Printf("ccpText=%d ccpFtn=%d ccpHdd=%d\n", ccpText, ccpFtn, ccpHdd)

	// Find fcPlcfHdd and lcbPlcfHdd
	// In FibRgFcLcb97, fcPlcfHdd is at offset 0x00F2 in the table
	// But we need to find it in the FIB structure
	// FibBase: 32 bytes (0x0000-0x001F)
	// csw: 2 bytes at 0x0020
	// FibRgW: csw*2 bytes
	// cslw: 2 bytes
	// FibRgLw: cslw*4 bytes
	// cbRgFcLcb: 2 bytes
	// FibRgFcLcbBlob: cbRgFcLcb*8 bytes

	csw := binary.LittleEndian.Uint16(wordDocData[0x0020:])
	fibRgWStart := uint32(0x0022)
	cslwOff := fibRgWStart + uint32(csw)*2
	cslw := binary.LittleEndian.Uint16(wordDocData[cslwOff:])
	fibRgLwStart := cslwOff + 2
	cbRgFcLcbOff := fibRgLwStart + uint32(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wordDocData[cbRgFcLcbOff:])
	fibRgFcLcbStart := cbRgFcLcbOff + 2

	fmt.Printf("csw=%d cslw=%d cbRgFcLcb=%d fibRgFcLcbStart=0x%X\n", csw, cslw, cbRgFcLcb, fibRgFcLcbStart)

	// fcPlcfHdd is at index 31 in FibRgFcLcb97 (0-based)
	// Each entry is 8 bytes (fc + lcb)
	hddIdx := uint32(31)
	fcPlcfHdd := binary.LittleEndian.Uint32(wordDocData[fibRgFcLcbStart+hddIdx*8:])
	lcbPlcfHdd := binary.LittleEndian.Uint32(wordDocData[fibRgFcLcbStart+hddIdx*8+4:])

	fmt.Printf("fcPlcfHdd=%d lcbPlcfHdd=%d\n", fcPlcfHdd, lcbPlcfHdd)

	if lcbPlcfHdd > 0 && uint64(fcPlcfHdd)+uint64(lcbPlcfHdd) <= uint64(len(tableData)) {
		plcData := tableData[fcPlcfHdd : fcPlcfHdd+lcbPlcfHdd]
		nCPs := lcbPlcfHdd / 4
		fmt.Printf("Number of CPs: %d (stories: %d)\n", nCPs, nCPs-1)

		cps := make([]uint32, nCPs)
		for i := uint32(0); i < nCPs; i++ {
			cps[i] = binary.LittleEndian.Uint32(plcData[i*4:])
		}

		hddStart := ccpText + ccpFtn
		hddEnd := hddStart + ccpHdd

		fmt.Printf("hddStart=%d hddEnd=%d\n", hddStart, hddEnd)

		for i := 0; i+1 < int(nCPs); i++ {
			cpStart := hddStart + cps[i]
			cpEnd := hddStart + cps[i+1]
			storyIdx := i % 6
			storyName := ""
			switch storyIdx {
			case 0:
				storyName = "even-header"
			case 1:
				storyName = "odd-header"
			case 2:
				storyName = "even-footer"
			case 3:
				storyName = "odd-footer"
			case 4:
				storyName = "first-header"
			case 5:
				storyName = "first-footer"
			}
			fmt.Printf("Story[%d] (%s): cp[%d-%d] len=%d\n", i, storyName, cpStart, cpEnd, cpEnd-cpStart)
		}
	}
}
