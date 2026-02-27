package main

import (
	"encoding/binary"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/helpers"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	fmt.Printf("Headers: %v\n", fc.Headers)
	fmt.Printf("Footers: %v\n", fc.Footers)
	for i, h := range fc.Headers {
		fmt.Printf("Header[%d] len=%d: %q\n", i, len(h), h)
		for _, r := range h {
			fmt.Printf("  U+%04X %c\n", r, r)
		}
	}
	for i, f := range fc.Footers {
		fmt.Printf("Footer[%d] len=%d: %q\n", i, len(f), f)
	}

	// Now let's look at the raw header/footer text
	rawText := d.GetText()
	runes := []rune(rawText)
	
	// Get FIB info by re-opening
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
	ccpText := binary.LittleEndian.Uint32(wordDocData[0x4C:0x50])
	ccpFtn := binary.LittleEndian.Uint32(wordDocData[0x50:0x54])
	ccpHdd := binary.LittleEndian.Uint32(wordDocData[0x54:0x58])
	
	fmt.Printf("\nccpText=%d ccpFtn=%d ccpHdd=%d\n", ccpText, ccpFtn, ccpHdd)
	fmt.Printf("Total runes: %d\n", len(runes))
	
	hddStart := ccpText + ccpFtn
	hddEnd := hddStart + ccpHdd
	fmt.Printf("Header/footer text range: rune %d to %d\n", hddStart, hddEnd)
	
	if int(hddEnd) <= len(runes) {
		hddText := string(runes[hddStart:hddEnd])
		fmt.Printf("\nRaw header/footer text (%d chars):\n", len([]rune(hddText)))
		for i, r := range []rune(hddText) {
			if r < 0x20 && r != '\t' {
				fmt.Printf("[%02X]", r)
			} else {
				fmt.Printf("%c", r)
			}
			if i > 500 {
				fmt.Println("\n... (truncated)")
				break
			}
		}
		fmt.Println()
	}
	
	// Also check the PlcfHdd
	// Need to find fcPlcfHdd from FIB
	// FIB base = 32 bytes, then csw, FibRgW, cslw, FibRgLw, cbRgFcLcb, FibRgFcLcb
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
	
	fcPlcfHdd := readFcLcb(22)
	lcbPlcfHdd := readFcLcb(23)
	fmt.Printf("\nfcPlcfHdd=%d lcbPlcfHdd=%d\n", fcPlcfHdd, lcbPlcfHdd)
	
	if lcbPlcfHdd > 0 && uint64(fcPlcfHdd)+uint64(lcbPlcfHdd) <= uint64(len(tableData)) {
		plcData := tableData[fcPlcfHdd : fcPlcfHdd+lcbPlcfHdd]
		nCPs := lcbPlcfHdd / 4
		fmt.Printf("PlcfHdd: %d CPs\n", nCPs)
		for i := uint32(0); i < nCPs; i++ {
			cp := binary.LittleEndian.Uint32(plcData[i*4:])
			absCP := hddStart + cp
			fmt.Printf("  CP[%d] = %d (abs=%d)", i, cp, absCP)
			if int(absCP) < len(runes) {
				// Show a few chars at this position
				end := int(absCP) + 20
				if end > len(runes) {
					end = len(runes)
				}
				snippet := string(runes[absCP:end])
				var display strings.Builder
				for _, r := range snippet {
					if r < 0x20 && r != '\t' {
						fmt.Fprintf(&display, "[%02X]", r)
					} else {
						display.WriteRune(r)
					}
				}
				fmt.Printf(" text=%q", display.String())
			}
			fmt.Println()
		}
	}
	
	_ = helpers.DecodeANSI // just to use the import
}
