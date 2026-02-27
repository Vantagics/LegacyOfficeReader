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

	var root, wordDoc, table1 *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "WordDocument":
			wordDoc = dir
		case "1Table":
			table1 = dir
		}
	}

	wReader, _ := adaptor.OpenObject(wordDoc, root)
	wSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wData := make([]byte, wSize)
	wReader.Read(wData)

	tReader, _ := adaptor.OpenObject(table1, root)
	tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tData := make([]byte, tSize)
	tReader.Read(tData)

	// Parse FIB
	offset := 0x20
	csw := binary.LittleEndian.Uint16(wData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wData[offset:])
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wData[offset:])
	offset += 2

	readFcLcb := func(index int) uint32 {
		if int(cbRgFcLcb) <= index {
			return 0
		}
		off := offset + index*4
		return binary.LittleEndian.Uint32(wData[off:])
	}

	fcPlcSpaMom := readFcLcb(82)
	lcbPlcSpaMom := readFcLcb(83)
	fmt.Printf("fcPlcSpaMom=%d lcbPlcSpaMom=%d\n", fcPlcSpaMom, lcbPlcSpaMom)

	if lcbPlcSpaMom == 0 {
		fmt.Println("No PlcSpaMom data")
		return
	}

	spaData := tData[fcPlcSpaMom : fcPlcSpaMom+lcbPlcSpaMom]

	// Dump raw bytes
	fmt.Printf("Raw PlcSpaMom bytes (%d):\n", lcbPlcSpaMom)
	for i := 0; i < int(lcbPlcSpaMom) && i < 200; i++ {
		fmt.Printf("%02X ", spaData[i])
		if (i+1)%16 == 0 {
			fmt.Println()
		}
	}
	fmt.Println()

	// PlcSpa structure: (n+1) CPs (uint32) + n SPAs (26 bytes each)
	// lcb = (n+1)*4 + n*26 = 4 + 30*n
	n := (lcbPlcSpaMom - 4) / 30
	fmt.Printf("n=%d entries\n", n)

	// Read CPs
	fmt.Println("\nCPs:")
	for i := uint32(0); i <= n; i++ {
		cp := binary.LittleEndian.Uint32(spaData[i*4:])
		fmt.Printf("  CP[%d] = %d\n", i, cp)
	}

	// Read SPAs
	fmt.Println("\nSPAs:")
	for i := uint32(0); i < n; i++ {
		spaOff := (n+1)*4 + i*26
		// SPA structure per MS-DOC 2.9.252:
		// uint32 spid
		// int32 xaLeft, yaTop, xaRight, yaBottom (position in twips)
		// uint16 flags
		// uint32 cTxbx (text box count, only for text boxes)
		spid := binary.LittleEndian.Uint32(spaData[spaOff:])
		xaLeft := int32(binary.LittleEndian.Uint32(spaData[spaOff+4:]))
		yaTop := int32(binary.LittleEndian.Uint32(spaData[spaOff+8:]))
		xaRight := int32(binary.LittleEndian.Uint32(spaData[spaOff+12:]))
		yaBottom := int32(binary.LittleEndian.Uint32(spaData[spaOff+16:]))
		flags := binary.LittleEndian.Uint16(spaData[spaOff+20:])
		cTxbx := binary.LittleEndian.Uint32(spaData[spaOff+22:])

		fmt.Printf("  SPA[%d]: spid=%d pos=(%d,%d)-(%d,%d) flags=0x%04X cTxbx=%d\n",
			i, spid, xaLeft, yaTop, xaRight, yaBottom, flags, cTxbx)

		// Decode flags
		fHdr := (flags >> 0) & 1
		bx := (flags >> 1) & 3
		by := (flags >> 3) & 3
		wr := (flags >> 5) & 0xF
		wrk := (flags >> 9) & 0xF
		fRcaSimple := (flags >> 13) & 1
		fBelowText := (flags >> 14) & 1
		fAnchorLock := (flags >> 15) & 1
		fmt.Printf("    fHdr=%d bx=%d by=%d wr=%d wrk=%d fRcaSimple=%d fBelowText=%d fAnchorLock=%d\n",
			fHdr, bx, by, wr, wrk, fRcaSimple, fBelowText, fAnchorLock)
	}
}
