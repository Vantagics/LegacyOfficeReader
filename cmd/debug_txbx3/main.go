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
	var root, wordDocDir, tableDir, dataDir *cfb.Directory
	for _, d := range dirs {
		name := d.Name()
		if name == "Root Entry" { root = d }
		if name == "WordDocument" { wordDocDir = d }
		if name == "1Table" || name == "0Table" {
			if tableDir == nil { tableDir = d }
		}
		if name == "Data" { dataDir = d }
	}

	wdReader, _ := c.OpenObject(wordDocDir, root)
	wordDocData, _ := io.ReadAll(wdReader)
	tReader, _ := c.OpenObject(tableDir, root)
	tableData, _ := io.ReadAll(tReader)

	var dataStreamBytes []byte
	if dataDir != nil {
		dReader, _ := c.OpenObject(dataDir, root)
		dataStreamBytes, _ = io.ReadAll(dReader)
		fmt.Printf("Data stream: %d bytes\n", len(dataStreamBytes))
	}

	// Parse PlcSpaMom to get shape CPs and SPIDs
	fcPlcSpaMom := binary.LittleEndian.Uint32(wordDocData[0x1DA:0x1DE])
	lcbPlcSpaMom := binary.LittleEndian.Uint32(wordDocData[0x1DE:0x1E2])
	
	spaData := tableData[fcPlcSpaMom : fcPlcSpaMom+lcbPlcSpaMom]
	nSpa := (lcbPlcSpaMom - 4) / 30
	
	fmt.Printf("SPAs (%d):\n", nSpa)
	for i := uint32(0); i < nSpa; i++ {
		cp := binary.LittleEndian.Uint32(spaData[i*4:])
		spaOff := (nSpa+1)*4 + i*26
		spid := binary.LittleEndian.Uint32(spaData[spaOff:])
		// SPA also has: xaLeft, yaTop, xaRight, yaBottom, flags
		xaLeft := int32(binary.LittleEndian.Uint32(spaData[spaOff+4:]))
		yaTop := int32(binary.LittleEndian.Uint32(spaData[spaOff+8:]))
		xaRight := int32(binary.LittleEndian.Uint32(spaData[spaOff+12:]))
		yaBottom := int32(binary.LittleEndian.Uint32(spaData[spaOff+16:]))
		flags := binary.LittleEndian.Uint16(spaData[spaOff+20:])
		fmt.Printf("  SPA[%d]: cp=%d spid=%d pos=(%d,%d)-(%d,%d) flags=0x%04X\n",
			i, cp, spid, xaLeft, yaTop, xaRight, yaBottom, flags)
	}

	// Now scan the Data stream for SpContainers and find text box shapes
	// Each shape referenced by PlcSpaMom has its SpContainer in the Data stream
	// The SpContainer offset is determined by the sprmCPicLocation in CHPX
	// But for drawn objects (0x08), the SPA directly references the shape by SPID
	
	// The shapes are in the DggContainer in the table stream
	// fcDggInfo points to the OfficeArt data in the table stream
	fcDggInfo := binary.LittleEndian.Uint32(wordDocData[0x1E2:0x1E6])
	lcbDggInfo := binary.LittleEndian.Uint32(wordDocData[0x1E6:0x1EA])
	fmt.Printf("\nDggInfo: fc=%d lcb=%d\n", fcDggInfo, lcbDggInfo)

	if lcbDggInfo > 0 {
		dggData := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]
		fmt.Printf("Parsing DggContainer (%d bytes)...\n", len(dggData))
		parseContainer(dggData, 0, len(dggData), 0)
	}

	_ = dataStreamBytes
}

func parseContainer(data []byte, offset, end, depth int) {
	indent := ""
	for i := 0; i < depth; i++ {
		indent += "  "
	}

	for offset+8 <= end {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])

		ver := verInst & 0x0F
		inst := verInst >> 4

		name := msoRecordName(recType)

		if ver == 0x0F { // container
			fmt.Printf("%s[0x%04X] %s inst=%d len=%d\n", indent, recType, name, inst, recLen)
			innerEnd := offset + 8 + int(recLen)
			if innerEnd > end {
				innerEnd = end
			}
			parseContainer(data, offset+8, innerEnd, depth+1)
			offset = innerEnd
		} else {
			fmt.Printf("%s[0x%04X] %s ver=%d inst=%d len=%d", indent, recType, name, ver, inst, recLen)
			
			recData := data[offset+8:]
			recEnd := int(recLen)
			if offset+8+recEnd > end {
				recEnd = end - offset - 8
			}
			if recEnd < 0 {
				recEnd = 0
			}
			recData = recData[:recEnd]

			if recType == 0xF00A && len(recData) >= 8 { // Sp
				spid := binary.LittleEndian.Uint32(recData[0:])
				flags := binary.LittleEndian.Uint32(recData[4:])
				fmt.Printf(" spid=%d flags=0x%08X", spid, flags)
			}
			
			if recType == 0xF00B || recType == 0xF121 || recType == 0xF122 { // OPT, SecondaryOPT, TertiaryOPT
				parseOPTRecord(recData, indent+"  ", inst)
			}

			fmt.Println()
			offset += 8 + int(recLen)
		}
	}
}

func parseOPTRecord(data []byte, indent string, numProps uint16) {
	fmt.Println()
	propTableSize := int(numProps) * 6
	if propTableSize > len(data) {
		propTableSize = len(data)
	}
	
	for i := 0; i < int(numProps); i++ {
		off := i * 6
		if off+6 > len(data) {
			break
		}
		propID := binary.LittleEndian.Uint16(data[off:])
		propVal := binary.LittleEndian.Uint32(data[off+2:])
		pid := propID & 0x3FFF
		isComplex := propID&0x8000 != 0
		_ = isComplex

		switch pid {
		case 0x0080: // lTxid
			fmt.Printf("%slTxid=%d\n", indent, propVal)
		case 0x0104: // pib
			fmt.Printf("%spib=%d\n", indent, propVal)
		case 0x0085: // pWrapPolygonVertices
			fmt.Printf("%spWrapPolygonVertices=%d\n", indent, propVal)
		}
	}
}

func msoRecordName(t uint16) string {
	switch t {
	case 0xF000: return "DggContainer"
	case 0xF001: return "BStoreContainer"
	case 0xF002: return "DgContainer"
	case 0xF003: return "SpgrContainer"
	case 0xF004: return "SpContainer"
	case 0xF006: return "Dgg"
	case 0xF007: return "BSE"
	case 0xF008: return "Dg"
	case 0xF009: return "Spgr"
	case 0xF00A: return "Sp"
	case 0xF00B: return "OPT"
	case 0xF00D: return "ClientTextbox"
	case 0xF010: return "ClientAnchor"
	case 0xF011: return "ClientData"
	case 0xF01E: return "SplitMenuColors"
	case 0xF11E: return "TertiaryOPT"
	case 0xF121: return "SecondaryOPT"
	case 0xF122: return "ColorScheme"
	default: return fmt.Sprintf("Unknown(0x%04X)", t)
	}
}
