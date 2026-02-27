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
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "WordDocument":
			wordDoc = dir
		}
	}

	reader, err := adaptor.OpenObject(wordDoc, root)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	size := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	data := make([]byte, size)
	reader.Read(data)

	fmt.Printf("WordDocument stream: %d bytes\n", len(data))

	// Parse FIB to find OfficeArt Drawing Layer offsets
	// FibRgFcLcb97 has fcDggInfo at index 80, lcbDggInfo at index 81
	// Navigate to FibRgFcLcb
	offset := 0x20
	csw := binary.LittleEndian.Uint16(data[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(data[offset:])
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(data[offset:])
	offset += 2

	readFcLcb := func(index int) uint32 {
		if int(cbRgFcLcb) <= index {
			return 0
		}
		off := offset + index*4
		if off+4 > len(data) {
			return 0
		}
		return binary.LittleEndian.Uint32(data[off:])
	}

	// OfficeArt Drawing Layer info
	fcDggInfo := readFcLcb(80)
	lcbDggInfo := readFcLcb(81)
	fmt.Printf("fcDggInfo=%d, lcbDggInfo=%d\n", fcDggInfo, lcbDggInfo)

	// Read the Table stream to get DggInfo
	var table1 *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		if dir.Name() == "1Table" {
			table1 = dir
		}
	}

	if table1 != nil && lcbDggInfo > 0 {
		tReader, _ := adaptor.OpenObject(table1, root)
		tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
		tData := make([]byte, tSize)
		tReader.Read(tData)

		if uint64(fcDggInfo)+uint64(lcbDggInfo) <= uint64(len(tData)) {
			dggData := tData[fcDggInfo : fcDggInfo+lcbDggInfo]
			fmt.Printf("DggInfo data: %d bytes\n", len(dggData))

			// Scan for OfficeArt records in DggInfo
			scanOARecords(dggData, 0, uint32(len(dggData)), 0)
		}
	}
}

func scanOARecords(data []byte, offset, limit uint32, depth int) {
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

		childStart := offset + 8
		childEnd := childStart + recLen
		if childEnd > limit {
			break
		}

		typeName := getTypeName(recType)
		fmt.Printf("%soffset=%d ver=0x%X inst=0x%03X type=0x%04X(%s) len=%d\n",
			indent, offset, recVer, recInst, recType, typeName, recLen)

		if recType == 0xF007 && recVer == 0x2 {
			// BSE record
			if offset+8+36 <= limit {
				btWin32 := data[offset+8]
				btMacOS := data[offset+9]
				fmt.Printf("%s  BSE btWin32=%d btMacOS=%d\n", indent, btWin32, btMacOS)
				// Check for embedded blip
				if recLen > 36 {
					blipOff := offset + 8 + 36
					if blipOff+8 <= limit {
						blipType := binary.LittleEndian.Uint16(data[blipOff+2 : blipOff+4])
						blipLen := binary.LittleEndian.Uint32(data[blipOff+4 : blipOff+8])
						fmt.Printf("%s  Embedded blip: type=0x%04X len=%d\n", indent, blipType, blipLen)
					}
				} else {
					fmt.Printf("%s  BSE references external blip (no embedded data)\n", indent)
				}
			}
		}

		if recVer == 0xF {
			scanOARecords(data, childStart, childEnd, depth+1)
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
	case 0xF002:
		return "DgContainer"
	case 0xF003:
		return "SpgrContainer"
	case 0xF004:
		return "SpContainer"
	case 0xF006:
		return "Dgg"
	case 0xF007:
		return "BSE"
	case 0xF008:
		return "Dg"
	case 0xF009:
		return "Spgr"
	case 0xF00A:
		return "Sp"
	case 0xF00B:
		return "Opt"
	case 0xF010:
		return "ClientAnchor"
	case 0xF011:
		return "ClientData"
	case 0xF01A:
		return "BlipEMF"
	case 0xF01B:
		return "BlipWMF"
	case 0xF01D:
		return "BlipJPEG"
	case 0xF01E:
		return "BlipPNG"
	case 0xF01F:
		return "BlipDIB"
	case 0xF11E:
		return "SplitMenuColors"
	case 0xF122:
		return "TertiaryOpt"
	default:
		return fmt.Sprintf("0x%04X", recType)
	}
}
