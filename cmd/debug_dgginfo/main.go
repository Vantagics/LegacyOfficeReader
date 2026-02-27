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

	var root, wordDoc *cfb.Directory
	var table1 *cfb.Directory
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

	// Read WordDocument stream
	wReader, _ := adaptor.OpenObject(wordDoc, root)
	wSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wData := make([]byte, wSize)
	wReader.Read(wData)

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
		if off+4 > len(wData) {
			return 0
		}
		return binary.LittleEndian.Uint32(wData[off:])
	}

	fcDggInfo := readFcLcb(80)
	lcbDggInfo := readFcLcb(81)
	fmt.Printf("fcDggInfo=%d, lcbDggInfo=%d\n", fcDggInfo, lcbDggInfo)

	// Also check for PlcSpaMom (main document shapes) and PlcSpaHdr (header shapes)
	// PlcSpaMom: index 82/83
	fcPlcSpaMom := readFcLcb(82)
	lcbPlcSpaMom := readFcLcb(83)
	fmt.Printf("fcPlcSpaMom=%d, lcbPlcSpaMom=%d\n", fcPlcSpaMom, lcbPlcSpaMom)

	// PlcSpaHdr: index 84/85
	fcPlcSpaHdr := readFcLcb(84)
	lcbPlcSpaHdr := readFcLcb(85)
	fmt.Printf("fcPlcSpaHdr=%d, lcbPlcSpaHdr=%d\n", fcPlcSpaHdr, lcbPlcSpaHdr)

	// Read Table stream
	tReader, _ := adaptor.OpenObject(table1, root)
	tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tData := make([]byte, tSize)
	tReader.Read(tData)

	// Parse DggInfo from Table stream
	if lcbDggInfo > 0 && uint64(fcDggInfo)+uint64(lcbDggInfo) <= uint64(len(tData)) {
		dggData := tData[fcDggInfo : fcDggInfo+lcbDggInfo]
		fmt.Printf("\n=== DggInfo (%d bytes) ===\n", len(dggData))
		// Hex dump first 100 bytes
		limit := 100
		if len(dggData) < limit {
			limit = len(dggData)
		}
		for i := 0; i < limit; i++ {
			if i%16 == 0 {
				fmt.Printf("\n%04X: ", i)
			}
			fmt.Printf("%02X ", dggData[i])
		}
		fmt.Println()

		// Try to parse as OfficeArt records
		fmt.Println("\n=== DggInfo OfficeArt Records ===")
		parseOARecords(dggData, 0, uint32(len(dggData)), 0)
	}

	// Parse PlcSpaMom (shape positions in main document)
	if lcbPlcSpaMom > 0 && uint64(fcPlcSpaMom)+uint64(lcbPlcSpaMom) <= uint64(len(tData)) {
		spaData := tData[fcPlcSpaMom : fcPlcSpaMom+lcbPlcSpaMom]
		fmt.Printf("\n=== PlcSpaMom (%d bytes) ===\n", len(spaData))
		// PlcSpaMom: (n+1) CPs (uint32) + n SPA entries (26 bytes each)
		// lcb = (n+1)*4 + n*26 = 4 + 30*n => n = (lcb - 4) / 30
		n := (lcbPlcSpaMom - 4) / 30
		fmt.Printf("Number of shapes: %d\n", n)
		for i := uint32(0); i <= n; i++ {
			cp := binary.LittleEndian.Uint32(spaData[i*4:])
			fmt.Printf("  CP[%d] = %d\n", i, cp)
		}
		// SPA entries start at (n+1)*4
		spaStart := (n + 1) * 4
		for i := uint32(0); i < n; i++ {
			off := spaStart + i*26
			if off+26 > lcbPlcSpaMom {
				break
			}
			spid := binary.LittleEndian.Uint32(spaData[off:])
			// SPA: spid(4) + xaLeft(4) + yaTop(4) + xaRight(4) + yaBottom(4) + flags(2) + padding(4)
			xaLeft := int32(binary.LittleEndian.Uint32(spaData[off+4:]))
			yaTop := int32(binary.LittleEndian.Uint32(spaData[off+8:]))
			xaRight := int32(binary.LittleEndian.Uint32(spaData[off+12:]))
			yaBottom := int32(binary.LittleEndian.Uint32(spaData[off+16:]))
			flags := binary.LittleEndian.Uint16(spaData[off+20:])
			fmt.Printf("  SPA[%d]: spid=%d rect=(%d,%d,%d,%d) flags=0x%04X\n",
				i, spid, xaLeft, yaTop, xaRight, yaBottom, flags)
		}
	}
}

func parseOARecords(data []byte, offset, limit uint32, depth int) {
	indent := ""
	for i := 0; i < depth; i++ {
		indent += "  "
	}

	for offset+8 <= limit {
		verInst := binary.LittleEndian.Uint16(data[offset : offset+2])
		recVer := verInst & 0x0F
		recInst := verInst >> 4
		recType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
		recLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])

		childEnd := offset + 8 + recLen
		if childEnd > limit {
			break
		}

		typeName := ""
		switch recType {
		case 0xF000:
			typeName = "DggContainer"
		case 0xF001:
			typeName = "BStoreContainer"
		case 0xF006:
			typeName = "Dgg"
		case 0xF007:
			typeName = "BSE"
		case 0xF00B:
			typeName = "Opt"
		case 0xF11E:
			typeName = "SplitMenuColors"
		case 0xF122:
			typeName = "TertiaryOpt"
		default:
			typeName = fmt.Sprintf("0x%04X", recType)
		}

		fmt.Printf("%soffset=%d ver=0x%X inst=0x%03X type=%s len=%d\n",
			indent, offset, recVer, recInst, typeName, recLen)

		if recType == 0xF007 {
			// BSE record
			if offset+8+36 <= limit {
				btWin32 := data[offset+8]
				fmt.Printf("%s  btWin32=%d (1=EMF,2=WMF,3=PICT,4=JPEG,5=PNG,6=DIB)\n", indent, btWin32)
				// foDelay at offset 8+24 (4 bytes) - offset into Data stream
				foDelay := binary.LittleEndian.Uint32(data[offset+8+24:])
				fmt.Printf("%s  foDelay=%d (offset in Data stream)\n", indent, foDelay)
				// cRef at offset 8+28 (4 bytes) - reference count
				cRef := binary.LittleEndian.Uint32(data[offset+8+28:])
				fmt.Printf("%s  cRef=%d (reference count)\n", indent, cRef)
			}
		}

		if recVer == 0xF {
			parseOARecords(data, offset+8, childEnd, depth+1)
		}

		offset = childEnd
	}
}
