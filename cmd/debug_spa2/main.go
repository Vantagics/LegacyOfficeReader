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

	// Check both PlcSpaMom (index 80) and PlcSpaHdr (index 82)
	fcPlcSpaMom := readFcLcb(80)
	lcbPlcSpaMom := readFcLcb(81)
	fcPlcSpaHdr := readFcLcb(82)
	lcbPlcSpaHdr := readFcLcb(83)

	fmt.Printf("PlcSpaMom: fc=%d lcb=%d\n", fcPlcSpaMom, lcbPlcSpaMom)
	fmt.Printf("PlcSpaHdr: fc=%d lcb=%d\n", fcPlcSpaHdr, lcbPlcSpaHdr)

	// Parse PlcSpaMom
	if lcbPlcSpaMom > 0 {
		spaData := tableData[fcPlcSpaMom : fcPlcSpaMom+lcbPlcSpaMom]
		nSpa := (lcbPlcSpaMom - 4) / 30
		fmt.Printf("\nPlcSpaMom SPAs (%d):\n", nSpa)
		for i := uint32(0); i < nSpa; i++ {
			cp := binary.LittleEndian.Uint32(spaData[i*4:])
			spaOff := (nSpa+1)*4 + i*26
			spid := binary.LittleEndian.Uint32(spaData[spaOff:])
			fmt.Printf("  SPA[%d]: cp=%d spid=%d\n", i, cp, spid)
		}
	}

	// Parse PlcSpaHdr
	if lcbPlcSpaHdr > 0 {
		spaData := tableData[fcPlcSpaHdr : fcPlcSpaHdr+lcbPlcSpaHdr]
		nSpa := (lcbPlcSpaHdr - 4) / 30
		fmt.Printf("\nPlcSpaHdr SPAs (%d):\n", nSpa)
		for i := uint32(0); i < nSpa; i++ {
			cp := binary.LittleEndian.Uint32(spaData[i*4:])
			spaOff := (nSpa+1)*4 + i*26
			spid := binary.LittleEndian.Uint32(spaData[spaOff:])
			fmt.Printf("  SPA[%d]: cp=%d spid=%d\n", i, cp, spid)
		}
	}
}
