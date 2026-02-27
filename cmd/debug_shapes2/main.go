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

	// Read WordDocument stream
	wReader, _ := adaptor.OpenObject(wordDoc, root)
	wSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wData := make([]byte, wSize)
	wReader.Read(wData)

	// Read Table stream
	tReader, _ := adaptor.OpenObject(table1, root)
	tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tData := make([]byte, tSize)
	tReader.Read(tData)

	// Parse FIB to get DggInfo offset
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

	// Also get PlcSpaMom (main document shapes)
	fcPlcSpaMom := readFcLcb(82)
	lcbPlcSpaMom := readFcLcb(83)
	fmt.Printf("fcPlcSpaMom=%d lcbPlcSpaMom=%d\n", fcPlcSpaMom, lcbPlcSpaMom)

	// Parse PlcSpaMom: (n+1) CPs + n SPAs
	if lcbPlcSpaMom > 0 {
		fmt.Println("\n=== PlcSpaMom (Main Document Shapes) ===")
		spaData := tData[fcPlcSpaMom : fcPlcSpaMom+lcbPlcSpaMom]
		// SPA is 26 bytes each
		// lcb = (n+1)*4 + n*26 => n = (lcb - 4) / 30
		n := (lcbPlcSpaMom - 4) / 30
		fmt.Printf("n=%d shapes\n", n)
		for i := uint32(0); i < n; i++ {
			cp := binary.LittleEndian.Uint32(spaData[i*4:])
			spaOff := (n+1)*4 + i*26
			spid := binary.LittleEndian.Uint32(spaData[spaOff:])
			fmt.Printf("  Shape %d: CP=%d SPID=%d\n", i, cp, spid)
		}
	}

	// Parse DggInfo from Table stream
	if lcbDggInfo > 0 {
		fmt.Println("\n=== DggInfo Structure ===")
		dggData := tData[fcDggInfo : fcDggInfo+lcbDggInfo]
		parseOAFull(dggData, 0, uint32(len(dggData)), 0)
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

		// BSE record details
		if recType == 0xF007 && offset+8+36 <= limit {
			btWin32 := data[offset+8]
			foDelay := binary.LittleEndian.Uint32(data[offset+8+28:])
			fmt.Printf("%s  btWin32=%d foDelay=%d\n", indent, btWin32, foDelay)
		}

		// SpContainer - parse children
		if recType == 0xF004 {
			parseSpContainer(data, offset+8, childEnd, depth+1)
		} else if recVer == 0xF {
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

		// Sp record (0xF00A) - shape info
		if recType == 0xF00A && recLen >= 8 {
			spid := binary.LittleEndian.Uint32(data[offset+8:])
			flags := binary.LittleEndian.Uint32(data[offset+12:])
			fmt.Printf("%s  SPID=%d flags=0x%08X\n", indent, spid, flags)
		}

		// Opt record (0xF00B) or TertiaryOpt (0xF122) - shape properties
		if (recType == 0xF00B || recType == 0xF122) && recLen > 0 {
			nProps := recInst
			fmt.Printf("%s  nProps=%d\n", indent, nProps)
			propOff := offset + 8
			for p := uint16(0); p < nProps && propOff+6 <= childEnd; p++ {
				propID := binary.LittleEndian.Uint16(data[propOff:])
				propVal := binary.LittleEndian.Uint32(data[propOff+2:])
				pid := propID & 0x3FFF
				fComplex := (propID >> 14) & 1
				fBid := (propID >> 15) & 1
				_ = fBid
				// Show all properties, highlight pib-related ones
				if pid == 260 || pid == 261 || pid == 262 { // pib, pibName, pibFlags
					fmt.Printf("%s  *** PROP pid=%d (pib-related) val=%d fComplex=%d fBid=%d\n",
						indent, pid, propVal, fComplex, fBid)
				} else if pid == 4 { // rotation
					fmt.Printf("%s  PROP pid=%d (rotation) val=%d\n", indent, pid, propVal)
				} else if pid >= 256 && pid <= 270 { // fill properties
					fmt.Printf("%s  PROP pid=%d (fill) val=%d fComplex=%d\n", indent, pid, propVal, fComplex)
				}
				propOff += 6
			}
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
