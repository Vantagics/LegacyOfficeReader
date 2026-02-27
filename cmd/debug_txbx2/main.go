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
	
	if lcbPlcSpaMom == 0 {
		fmt.Println("No PlcSpaMom")
		return
	}

	spaData := tableData[fcPlcSpaMom : fcPlcSpaMom+lcbPlcSpaMom]
	nSpa := (lcbPlcSpaMom - 4) / 30
	
	type spaEntry struct {
		cp   uint32
		spid uint32
	}
	var spas []spaEntry
	for i := uint32(0); i < nSpa; i++ {
		cp := binary.LittleEndian.Uint32(spaData[i*4:])
		spaOff := (nSpa+1)*4 + i*26
		spid := binary.LittleEndian.Uint32(spaData[spaOff:])
		spas = append(spas, spaEntry{cp, spid})
	}

	// Now parse the DggContainer in the Data stream to find shapes with txid
	// The DggContainer is referenced by fcDggInfo in the table stream
	fcDggInfo := binary.LittleEndian.Uint32(wordDocData[0x1E2:0x1E6])
	lcbDggInfo := binary.LittleEndian.Uint32(wordDocData[0x1E6:0x1EA])
	fmt.Printf("DggInfo: fc=%d lcb=%d\n", fcDggInfo, lcbDggInfo)

	if lcbDggInfo == 0 {
		return
	}

	dggData := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]
	fmt.Printf("DggInfo data: %d bytes\n", len(dggData))

	// Parse the OfficeArt DggContainer
	// It starts with a record header: ver(4bits) + inst(12bits) + type(16bits) + length(32bits)
	parseMSODrawing(dggData, 0, len(dggData), 0)

	// Also check shapes in the Data stream for text box references
	if len(dataStreamBytes) > 0 {
		fmt.Printf("\n=== Scanning Data stream for SpContainers ===\n")
		scanForTextBoxShapes(dataStreamBytes)
	}

	_ = spas
}

func parseMSODrawing(data []byte, offset, end, depth int) {
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
		fmt.Printf("%s[0x%04X] %s ver=%d inst=%d len=%d\n", indent, recType, name, ver, inst, recLen)

		if ver == 0x0F { // container
			parseMSODrawing(data, offset+8, offset+8+int(recLen), depth+1)
		} else {
			// For SpContainer children, look for interesting properties
			if recType == 0xF00B { // OPT record
				parseOPT(data[offset+8:offset+8+int(recLen)], indent+"  ", inst)
			}
			if recType == 0xF00A { // Sp record
				if recLen >= 8 {
					spid := binary.LittleEndian.Uint32(data[offset+8:])
					flags := binary.LittleEndian.Uint32(data[offset+12:])
					fmt.Printf("%s  spid=%d flags=0x%08X\n", indent, spid, flags)
				}
			}
		}

		offset += 8 + int(recLen)
	}
}

func parseOPT(data []byte, indent string, numProps uint16) {
	if len(data) < int(numProps)*6 {
		return
	}
	for i := uint16(0); i < numProps; i++ {
		off := int(i) * 6
		if off+6 > len(data) {
			break
		}
		propID := binary.LittleEndian.Uint16(data[off:])
		propVal := binary.LittleEndian.Uint32(data[off+2:])
		pid := propID & 0x3FFF
		isComplex := propID&0x8000 != 0
		isBlip := propID&0x4000 != 0

		// Show interesting properties
		switch pid {
		case 0x0080: // lTxid - text box ID
			fmt.Printf("%s  lTxid=%d\n", indent, propVal)
		case 0x0081: // dxTextLeft
			fmt.Printf("%s  dxTextLeft=%d\n", indent, propVal)
		case 0x0104: // pib (picture BSE index)
			fmt.Printf("%s  pib=%d\n", indent, propVal)
		case 0x0180: // geoLeft
		case 0x0181: // geoTop
		default:
			if pid >= 0x0080 && pid <= 0x0085 {
				fmt.Printf("%s  prop[0x%04X]=%d complex=%v blip=%v\n", indent, pid, propVal, isComplex, isBlip)
			}
		}
	}
}

func scanForTextBoxShapes(data []byte) {
	offset := 0
	for offset+8 <= len(data) {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])

		ver := verInst & 0x0F

		if recType == 0xF004 { // SpContainer
			fmt.Printf("SpContainer at offset %d, len=%d\n", offset, recLen)
			// Parse inside
			parseMSODrawing(data, offset+8, offset+8+int(recLen), 1)
		}

		if ver == 0x0F {
			offset += 8 // enter container
		} else {
			offset += 8 + int(recLen)
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
	case 0xF122: return "ColorScheme"
	default: return fmt.Sprintf("Unknown(0x%04X)", t)
	}
}
