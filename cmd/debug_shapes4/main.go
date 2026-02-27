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

	var root, wordDoc, table1, dataDir *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "WordDocument":
			wordDoc = dir
		case "1Table":
			table1 = dir
		case "Data":
			dataDir = dir
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

	var dData []byte
	if dataDir != nil {
		dReader, _ := adaptor.OpenObject(dataDir, root)
		dSize := binary.LittleEndian.Uint32(dataDir.StreamSize[:])
		dData = make([]byte, dSize)
		dReader.Read(dData)
	}

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

	fcDggInfo := readFcLcb(80)
	lcbDggInfo := readFcLcb(81)

	// Per MS-DOC 2.9.63, fcDggInfo points to OfficeArtContent in the Table stream.
	// OfficeArtContent = OfficeArtDggContainer + OfficeArtDgContainer(s)
	// But the first bytes don't look like OfficeArt records.
	// Let me dump more bytes and look for the DggContainer signature (0xF000)
	dggData := tData[fcDggInfo : fcDggInfo+lcbDggInfo]
	fmt.Printf("DggInfo: offset=%d len=%d\n", fcDggInfo, lcbDggInfo)
	fmt.Printf("First 64 bytes:\n")
	for i := 0; i < 64 && i < len(dggData); i++ {
		fmt.Printf("%02X ", dggData[i])
		if (i+1)%16 == 0 {
			fmt.Println()
		}
	}
	fmt.Println()

	// The first 4 bytes might be a count or offset. Let me check if there's
	// an OfficeArt record starting at various offsets
	for i := 0; i+8 <= len(dggData); i++ {
		recType := binary.LittleEndian.Uint16(dggData[i+2 : i+4])
		if recType == 0xF000 || recType == 0xF002 {
			verInst := binary.LittleEndian.Uint16(dggData[i : i+2])
			recVer := verInst & 0x0F
			recLen := binary.LittleEndian.Uint32(dggData[i+4 : i+8])
			if recVer == 0xF && recLen > 0 && uint32(i)+8+recLen <= uint32(len(dggData)) {
				fmt.Printf("Found %s at DggInfo offset %d, len=%d\n", getTypeName(recType), i, recLen)
			}
		}
	}

	// Now let me understand the Data stream structure better.
	// The shapes in PlcSpaMom reference SPIDs 2049-2051.
	// The Data stream has SPIDs 1025-1027.
	// Per MS-DOC, the shapes for the main document are stored in the Data stream
	// as OfficeArtWordDrawing structures.
	// Let me scan the Data stream for all OfficeArt containers
	fmt.Println("\n=== Data stream OfficeArt containers ===")
	for i := 0; i+8 <= len(dData); i++ {
		verInst := binary.LittleEndian.Uint16(dData[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(dData[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(dData[i+4 : i+8])

		if recVer == 0xF && recType >= 0xF000 && recType <= 0xF00F && recLen > 0 && uint32(i)+8+recLen <= uint32(len(dData)) {
			fmt.Printf("  offset=%d type=%s len=%d\n", i, getTypeName(recType), recLen)
			if recType == 0xF002 || recType == 0xF003 || recType == 0xF004 {
				// Parse children to find shapes
				parseShapes(dData, uint32(i)+8, uint32(i)+8+recLen, 2)
			}
		}
	}

	// Check: what's at the beginning of the Data stream?
	fmt.Println("\n=== Data stream first 64 bytes ===")
	for i := 0; i < 64 && i < len(dData); i++ {
		fmt.Printf("%02X ", dData[i])
		if (i+1)%16 == 0 {
			fmt.Println()
		}
	}
	fmt.Println()

	// Per MS-DOC, the Data stream contains OfficeArtWordDrawing structures.
	// Each starts with a DgContainer (0xF002).
	// Let me find all DgContainers
	fmt.Println("\n=== DgContainers in Data stream ===")
	for i := 0; i+8 <= len(dData); i++ {
		verInst := binary.LittleEndian.Uint16(dData[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(dData[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(dData[i+4 : i+8])

		if recType == 0xF002 && recVer == 0xF && recLen > 0 && uint32(i)+8+recLen <= uint32(len(dData)) {
			fmt.Printf("DgContainer at offset %d, len=%d\n", i, recLen)
			parseShapes(dData, uint32(i)+8, uint32(i)+8+recLen, 1)
		}
	}
}

func parseShapes(data []byte, offset, limit uint32, depth int) {
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
		fmt.Printf("%s[%d] type=%s inst=0x%03X len=%d\n", indent, offset, typeName, recInst, recLen)

		if recType == 0xF00A && recLen >= 8 {
			spid := binary.LittleEndian.Uint32(data[offset+8:])
			flags := binary.LittleEndian.Uint32(data[offset+12:])
			fmt.Printf("%s  SPID=%d flags=0x%08X\n", indent, spid, flags)
		}

		if (recType == 0xF00B || recType == 0xF122) && recLen > 0 {
			nProps := recInst
			propOff := offset + 8
			for p := uint16(0); p < nProps && propOff+6 <= childEnd; p++ {
				propID := binary.LittleEndian.Uint16(data[propOff:])
				propVal := binary.LittleEndian.Uint32(data[propOff+2:])
				pid := propID & 0x3FFF
				if pid == 260 {
					fmt.Printf("%s  *** pib=%d\n", indent, propVal)
				}
				propOff += 6
			}
		}

		if recType == 0xF008 && recLen >= 8 {
			numShapes := binary.LittleEndian.Uint32(data[offset+8:])
			lastSpid := binary.LittleEndian.Uint32(data[offset+12:])
			fmt.Printf("%s  numShapes=%d lastSpid=%d\n", indent, numShapes, lastSpid)
		}

		if recVer == 0xF {
			parseShapes(data, offset+8, childEnd, depth+1)
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
	case 0xF00D:
		return "ClientTextbox"
	case 0xF010:
		return "ClientAnchor"
	case 0xF011:
		return "ClientData"
	case 0xF11E:
		return "SplitMenuColors"
	case 0xF122:
		return "TertiaryOpt"
	default:
		if recType >= 0xF01A && recType <= 0xF02A {
			return fmt.Sprintf("Blip(0x%04X)", recType)
		}
		return fmt.Sprintf("0x%04X", recType)
	}
}
