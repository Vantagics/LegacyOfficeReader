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

	fcDggInfo := readFcLcb(100)
	lcbDggInfo := readFcLcb(101)
	fmt.Printf("fcDggInfo=%d lcbDggInfo=%d\n", fcDggInfo, lcbDggInfo)

	if lcbDggInfo == 0 {
		return
	}

	dggData := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]
	fmt.Printf("Parsing DggContainer (%d bytes)...\n\n", len(dggData))
	parseContainer(dggData, 0, len(dggData), 0)
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
		childEnd := offset + 8 + int(recLen)
		if childEnd > end {
			childEnd = end
		}

		if ver == 0x0F { // container
			fmt.Printf("%s[0x%04X] %s inst=%d len=%d\n", indent, recType, name, inst, recLen)
			parseContainer(data, offset+8, childEnd, depth+1)
		} else {
			recData := data[offset+8 : childEnd]

			if recType == 0xF00A && len(recData) >= 8 { // Sp
				spid := binary.LittleEndian.Uint32(recData[0:])
				flags := binary.LittleEndian.Uint32(recData[4:])
				fmt.Printf("%s[Sp] spid=%d flags=0x%08X\n", indent, spid, flags)
			} else if recType == 0xF00B { // OPT
				fmt.Printf("%s[OPT] inst=%d\n", indent, inst)
				parseOPT(recData, indent+"  ", inst)
			} else if recType == 0xF11E { // TertiaryOPT
				fmt.Printf("%s[TertiaryOPT] inst=%d\n", indent, inst)
				parseOPT(recData, indent+"  ", inst)
			} else if recType == 0xF006 && len(recData) >= 16 { // Dgg
				spidMax := binary.LittleEndian.Uint32(recData[0:])
				cidcl := binary.LittleEndian.Uint32(recData[4:])
				cspSaved := binary.LittleEndian.Uint32(recData[8:])
				cdgSaved := binary.LittleEndian.Uint32(recData[12:])
				fmt.Printf("%s[Dgg] spidMax=%d cidcl=%d cspSaved=%d cdgSaved=%d\n",
					indent, spidMax, cidcl, cspSaved, cdgSaved)
			} else if recType == 0xF008 && len(recData) >= 8 { // Dg
				csp := binary.LittleEndian.Uint32(recData[0:])
				spidCur := binary.LittleEndian.Uint32(recData[4:])
				fmt.Printf("%s[Dg] csp=%d spidCur=%d\n", indent, csp, spidCur)
			} else if recType == 0xF00D { // ClientTextbox
				fmt.Printf("%s[ClientTextbox] len=%d\n", indent, recLen)
				if len(recData) >= 4 {
					txbxIndex := binary.LittleEndian.Uint32(recData[0:])
					fmt.Printf("%s  txbxIndex=%d\n", indent, txbxIndex)
				}
			} else {
				fmt.Printf("%s[0x%04X] %s len=%d\n", indent, recType, name, recLen)
			}
		}

		offset = childEnd
	}
}

func parseOPT(data []byte, indent string, numProps uint16) {
	for i := 0; i < int(numProps); i++ {
		off := i * 6
		if off+6 > len(data) {
			break
		}
		propID := binary.LittleEndian.Uint16(data[off:])
		propVal := binary.LittleEndian.Uint32(data[off+2:])
		pid := propID & 0x3FFF
		isComplex := propID&0x8000 != 0

		switch pid {
		case 0x0080: // lTxid
			fmt.Printf("%slTxid=%d\n", indent, propVal)
		case 0x0104: // pib
			fmt.Printf("%spib=%d\n", indent, propVal)
		case 0x0085: // pWrapPolygonVertices
			fmt.Printf("%spWrapPolygonVertices=%d complex=%v\n", indent, propVal, isComplex)
		case 0x007F: // fLockText etc
			fmt.Printf("%sfLockText/flags=0x%08X\n", indent, propVal)
		default:
			if pid < 0x0200 {
				fmt.Printf("%sprop[0x%04X]=%d complex=%v\n", indent, pid, propVal, isComplex)
			}
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
	default: return fmt.Sprintf("Unknown(0x%04X)", t)
	}
}
