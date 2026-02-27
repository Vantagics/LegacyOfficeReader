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

	var root *cfb.Directory
	var table1 *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "1Table":
			table1 = dir
		}
	}

	tReader, _ := adaptor.OpenObject(table1, root)
	tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tData := make([]byte, tSize)
	tReader.Read(tData)

	// Scan the entire Table stream for OfficeArt records
	fmt.Println("=== Scanning Table stream for OfficeArt records ===")
	for i := 0; i+8 <= len(tData); i++ {
		verInst := binary.LittleEndian.Uint16(tData[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(tData[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(tData[i+4 : i+8])

		// Look for DggContainer (0xF000) or BStoreContainer (0xF001)
		if recType == 0xF000 && recVer == 0xF {
			if uint32(i)+8+recLen <= uint32(len(tData)) {
				fmt.Printf("Found DggContainer at offset %d, len=%d\n", i, recLen)
				parseOA(tData, uint32(i), uint32(i)+8+recLen, 0)
			}
		}
		if recType == 0xF001 && recVer == 0xF {
			if uint32(i)+8+recLen <= uint32(len(tData)) {
				fmt.Printf("Found BStoreContainer at offset %d, len=%d\n", i, recLen)
				parseOA(tData, uint32(i), uint32(i)+8+recLen, 0)
			}
		}
	}
}

func parseOA(data []byte, offset, limit uint32, depth int) {
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

		typeName := getTypeName(recType)
		fmt.Printf("%s  offset=%d ver=0x%X inst=0x%03X type=%s len=%d\n",
			indent, offset, recVer, recInst, typeName, recLen)

		if recType == 0xF007 {
			if offset+8+36 <= limit {
				btWin32 := data[offset+8]
				foDelay := binary.LittleEndian.Uint32(data[offset+8+24:])
				cRef := binary.LittleEndian.Uint32(data[offset+8+28:])
				cbBlip := binary.LittleEndian.Uint32(data[offset+8+32:])
				fmt.Printf("%s    btWin32=%d foDelay=%d cRef=%d cbBlip=%d\n",
					indent, btWin32, foDelay, cRef, cbBlip)
			}
		}

		if recVer == 0xF {
			parseOA(data, offset+8, childEnd, depth+1)
		}

		offset = childEnd
	}
}

func getTypeName(recType uint16) string {
	switch recType {
	case 0xF000:
		return "DggContainer"
	case 0xF001:
		return "BStoreContainer"
	case 0xF006:
		return "Dgg"
	case 0xF007:
		return "BSE"
	case 0xF00B:
		return "Opt"
	case 0xF11E:
		return "SplitMenuColors"
	case 0xF122:
		return "TertiaryOpt"
	default:
		return fmt.Sprintf("0x%04X", recType)
	}
}
