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

	// Read streams
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
		fmt.Printf("Data stream size: %d\n", len(dData))
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
	fmt.Printf("fcDggInfo=%d lcbDggInfo=%d\n", fcDggInfo, lcbDggInfo)

	// The DggInfo in the Table stream is the OfficeArt Drawing Group data.
	// Per MS-DOC, it contains: OfficeArtDggContainer + OfficeArtDgContainer(s)
	// Let's scan for OfficeArt record headers within this range
	dggData := tData[fcDggInfo : fcDggInfo+lcbDggInfo]
	fmt.Printf("First 32 bytes of DggInfo: ")
	for i := 0; i < 32 && i < len(dggData); i++ {
		fmt.Printf("%02X ", dggData[i])
	}
	fmt.Println()

	// Scan for OfficeArt container records
	fmt.Println("\n=== Scanning DggInfo for OfficeArt records ===")
	for i := 0; i+8 <= len(dggData); i++ {
		verInst := binary.LittleEndian.Uint16(dggData[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(dggData[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(dggData[i+4 : i+8])

		// Only look for top-level containers
		if recVer == 0xF && recType >= 0xF000 && recType <= 0xF00F && recLen > 0 && uint32(i)+8+recLen <= uint32(len(dggData)) {
			fmt.Printf("Found %s at offset %d (abs %d), len=%d\n", getTypeName(recType), i, fcDggInfo+uint32(i), recLen)
			parseOAFull(dggData, uint32(i)+8, uint32(i)+8+recLen, 1)
		}
	}

	// Also parse Data stream for SpContainers
	if len(dData) > 0 {
		fmt.Println("\n=== Scanning Data stream for SpContainers ===")
		for i := 0; i+8 <= len(dData); i++ {
			verInst := binary.LittleEndian.Uint16(dData[i : i+2])
			recVer := verInst & 0x0F
			recType := binary.LittleEndian.Uint16(dData[i+2 : i+4])
			recLen := binary.LittleEndian.Uint32(dData[i+4 : i+8])

			if recType == 0xF004 && recVer == 0xF && recLen > 0 && uint32(i)+8+recLen <= uint32(len(dData)) {
				fmt.Printf("Found SpContainer at offset %d, len=%d\n", i, recLen)
				parseSpContainer(dData, uint32(i)+8, uint32(i)+8+recLen, 1)
			}
		}
	}

	// PlcSpaMom
	fcPlcSpaMom := readFcLcb(82)
	lcbPlcSpaMom := readFcLcb(83)
	if lcbPlcSpaMom > 0 {
		fmt.Println("\n=== PlcSpaMom ===")
		spaData := tData[fcPlcSpaMom : fcPlcSpaMom+lcbPlcSpaMom]
		n := (lcbPlcSpaMom - 4) / 30
		for i := uint32(0); i < n; i++ {
			cp := binary.LittleEndian.Uint32(spaData[i*4:])
			spaOff := (n+1)*4 + i*26
			spid := binary.LittleEndian.Uint32(spaData[spaOff:])
			// SPA also has position info
			xaLeft := int32(binary.LittleEndian.Uint32(spaData[spaOff+4:]))
			yaTop := int32(binary.LittleEndian.Uint32(spaData[spaOff+8:]))
			xaRight := int32(binary.LittleEndian.Uint32(spaData[spaOff+12:]))
			yaBottom := int32(binary.LittleEndian.Uint32(spaData[spaOff+16:]))
			flags := binary.LittleEndian.Uint16(spaData[spaOff+20:])
			fmt.Printf("  Shape %d: CP=%d SPID=%d pos=(%d,%d)-(%d,%d) flags=0x%04X\n",
				i, cp, spid, xaLeft, yaTop, xaRight, yaBottom, flags)
		}
	}
}

func parseOAFull(data []byte, offset, limit uint32, depth int) {
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
		fmt.Printf("%s[%d] ver=0x%X inst=0x%03X type=%s len=%d\n",
			indent, offset, recVer, recInst, typeName, recLen)

		// BSE record
		if recType == 0xF007 && offset+8+36 <= limit {
			btWin32 := data[offset+8]
			foDelay := binary.LittleEndian.Uint32(data[offset+8+28:])
			cbBlip := binary.LittleEndian.Uint32(data[offset+8+32:])
			fmt.Printf("%s  btWin32=%d foDelay=%d cbBlip=%d\n", indent, btWin32, foDelay, cbBlip)
		}

		// Sp record
		if recType == 0xF00A && recLen >= 8 {
			spid := binary.LittleEndian.Uint32(data[offset+8:])
			flags := binary.LittleEndian.Uint32(data[offset+12:])
			fmt.Printf("%s  SPID=%d flags=0x%08X\n", indent, spid, flags)
		}

		// Opt record - parse properties
		if (recType == 0xF00B || recType == 0xF122) && recLen > 0 {
			nProps := recInst
			propOff := offset + 8
			for p := uint16(0); p < nProps && propOff+6 <= childEnd; p++ {
				propID := binary.LittleEndian.Uint16(data[propOff:])
				propVal := binary.LittleEndian.Uint32(data[propOff+2:])
				pid := propID & 0x3FFF
				fBid := (propID >> 15) & 1
				if pid == 260 { // pib
					fmt.Printf("%s  *** pib=%d (BSE index, 1-based) fBid=%d\n", indent, propVal, fBid)
				} else if pid == 261 { // pibName
					fmt.Printf("%s  pibName val=%d\n", indent, propVal)
				} else if pid >= 256 && pid <= 270 {
					fmt.Printf("%s  fill pid=%d val=%d\n", indent, pid, propVal)
				}
				propOff += 6
			}
		}

		if recVer == 0xF {
			parseOAFull(data, offset+8, childEnd, depth+1)
		}

		offset = childEnd
	}
}

func parseSpContainer(data []byte, offset, limit uint32, depth int) {
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
		fmt.Printf("%s[%d] ver=0x%X inst=0x%03X type=%s len=%d\n",
			indent, offset, recVer, recInst, typeName, recLen)

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
				fBid := (propID >> 15) & 1
				if pid == 260 {
					fmt.Printf("%s  *** pib=%d (BSE index, 1-based) fBid=%d\n", indent, propVal, fBid)
				} else if pid >= 256 && pid <= 270 {
					fmt.Printf("%s  fill pid=%d val=%d\n", indent, pid, propVal)
				}
				propOff += 6
			}
		}

		// Check for blip records inside SpContainer
		if recType >= 0xF01A && recType <= 0xF02A {
			fmt.Printf("%s  (Blip record)\n", indent)
		}

		if recVer == 0xF {
			parseSpContainer(data, offset+8, childEnd, depth+1)
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
	case 0xF005:
		return "SolverContainer"
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
