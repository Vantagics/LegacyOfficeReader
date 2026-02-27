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

	// Navigate FIB
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

	// Corrected indices
	fcPlcSpaMom := readFcLcb(80)
	lcbPlcSpaMom := readFcLcb(81)
	fcDggInfo := readFcLcb(100)
	lcbDggInfo := readFcLcb(101)
	fcPlcftxbxTxt := readFcLcb(112)
	lcbPlcftxbxTxt := readFcLcb(113)

	fmt.Printf("fcPlcSpaMom=%d lcbPlcSpaMom=%d\n", fcPlcSpaMom, lcbPlcSpaMom)
	fmt.Printf("fcDggInfo=%d lcbDggInfo=%d\n", fcDggInfo, lcbDggInfo)
	fmt.Printf("fcPlcftxbxTxt=%d lcbPlcftxbxTxt=%d\n", fcPlcftxbxTxt, lcbPlcftxbxTxt)

	if lcbDggInfo > 0 && uint64(fcDggInfo)+uint64(lcbDggInfo) <= uint64(len(tableData)) {
		dggData := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]
		fmt.Printf("\nDggInfo first 32 bytes:\n")
		for i := 0; i < 32 && i < len(dggData); i++ {
			fmt.Printf("%02X ", dggData[i])
			if (i+1)%16 == 0 {
				fmt.Println()
			}
		}
		fmt.Println()

		// Check if it starts with a valid OfficeArt container
		if len(dggData) >= 8 {
			verInst := binary.LittleEndian.Uint16(dggData[0:])
			recType := binary.LittleEndian.Uint16(dggData[2:])
			recLen := binary.LittleEndian.Uint32(dggData[4:])
			ver := verInst & 0x0F
			fmt.Printf("First record: type=0x%04X ver=%d len=%d\n", recType, ver, recLen)
			if recType == 0xF000 {
				fmt.Println("Valid DggContainer!")
			}
		}
	}

	if lcbPlcftxbxTxt > 0 {
		fmt.Printf("\nPlcftxbxTxt data (%d bytes):\n", lcbPlcftxbxTxt)
		txbxData := tableData[fcPlcftxbxTxt : fcPlcftxbxTxt+lcbPlcftxbxTxt]
		for i := 0; i < len(txbxData) && i < 64; i++ {
			fmt.Printf("%02X ", txbxData[i])
			if (i+1)%16 == 0 {
				fmt.Println()
			}
		}
		fmt.Println()
	}
}
