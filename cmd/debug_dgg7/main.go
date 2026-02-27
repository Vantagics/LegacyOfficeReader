package main

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"

	"github.com/shakinm/xlsReader/cfb"
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

	offset := 0x20
	csw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2

	readFcLcb := func(index int) uint32 {
		if int(cbRgFcLcb) <= index { return 0 }
		off := offset + index*4
		if off+4 > len(wordDocData) { return 0 }
		return binary.LittleEndian.Uint32(wordDocData[off:])
	}

	fcDggInfo := readFcLcb(100)
	lcbDggInfo := readFcLcb(101)
	dggData := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]

	// Show bytes around offset 500
	fmt.Printf("Bytes at offset 496-520:\n")
	for i := 496; i < 520 && i < len(dggData); i++ {
		fmt.Printf("%02X ", dggData[i])
		if (i-496+1)%16 == 0 {
			fmt.Println()
		}
	}
	fmt.Println()

	// The DggContainer ends at offset 500 (8 + 492)
	// Check if the next record is actually a DgContainer
	// In Word, the OfficeArtContent wraps everything in a single structure
	// Let me try parsing from offset 500 as a DgContainer
	pos := 500
	if pos+8 <= len(dggData) {
		verInst := binary.LittleEndian.Uint16(dggData[pos:])
		recType := binary.LittleEndian.Uint16(dggData[pos+2:])
		recLen := binary.LittleEndian.Uint32(dggData[pos+4:])
		ver := verInst & 0x0F
		inst := verInst >> 4
		fmt.Printf("At offset 500: verInst=0x%04X type=0x%04X len=%d ver=%d inst=%d\n",
			verInst, recType, recLen, ver, inst)
		
		// 0x0201 with ver=0xF would be a DgContainer... but 0x0200 is not
		// Wait, let me check: DgContainer is 0xF002
		// verInst for DgContainer: ver=0xF, inst=0 → verInst = 0x000F
		// recType = 0xF002
		// So bytes would be: 0F 00 02 F0 ...
		
		// What we have: let me show the actual bytes
		fmt.Printf("Raw bytes: %02X %02X %02X %02X %02X %02X %02X %02X\n",
			dggData[pos], dggData[pos+1], dggData[pos+2], dggData[pos+3],
			dggData[pos+4], dggData[pos+5], dggData[pos+6], dggData[pos+7])
	}

	// Maybe the OfficeArtContent in Word is structured differently
	// Let me scan for 0xF002 (DgContainer) in the data
	fmt.Printf("\nScanning for DgContainer (0xF002) records...\n")
	for i := 0; i+8 <= len(dggData); i++ {
		recType := binary.LittleEndian.Uint16(dggData[i+2 : i+4])
		verInst := binary.LittleEndian.Uint16(dggData[i : i+2])
		ver := verInst & 0x0F
		if recType == 0xF002 && ver == 0x0F {
			recLen := binary.LittleEndian.Uint32(dggData[i+4 : i+8])
			fmt.Printf("  Found DgContainer at offset %d, len=%d\n", i, recLen)
		}
	}

	// Also scan for SpContainer (0xF004)
	fmt.Printf("\nScanning for SpContainer (0xF004) records...\n")
	count := 0
	for i := 0; i+8 <= len(dggData); i++ {
		recType := binary.LittleEndian.Uint16(dggData[i+2 : i+4])
		verInst := binary.LittleEndian.Uint16(dggData[i : i+2])
		ver := verInst & 0x0F
		if recType == 0xF004 && ver == 0x0F {
			recLen := binary.LittleEndian.Uint32(dggData[i+4 : i+8])
			if recLen < 10000 && recLen > 8 { // reasonable size
				fmt.Printf("  Found SpContainer at offset %d, len=%d\n", i, recLen)
				count++
				if count > 30 { break }
			}
		}
	}
}
