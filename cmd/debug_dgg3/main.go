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

	fcDggInfo := readFcLcb(80)
	lcbDggInfo := readFcLcb(81)
	fmt.Printf("fcDggInfo=%d lcbDggInfo=%d\n", fcDggInfo, lcbDggInfo)

	if lcbDggInfo == 0 {
		fmt.Println("No DggInfo")
		return
	}

	dggData := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]
	
	// Show first 64 bytes as hex
	fmt.Printf("First 64 bytes of DggInfo:\n")
	for i := 0; i < 64 && i < len(dggData); i++ {
		fmt.Printf("%02X ", dggData[i])
		if (i+1)%16 == 0 {
			fmt.Println()
		}
	}
	fmt.Println()

	// Try parsing as OfficeArt records
	fmt.Printf("\nParsing as OfficeArt records:\n")
	pos := 0
	for pos+8 <= len(dggData) {
		verInst := binary.LittleEndian.Uint16(dggData[pos:])
		recType := binary.LittleEndian.Uint16(dggData[pos+2:])
		recLen := binary.LittleEndian.Uint32(dggData[pos+4:])
		ver := verInst & 0x0F
		inst := verInst >> 4

		fmt.Printf("  offset=%d verInst=0x%04X recType=0x%04X recLen=%d ver=%d inst=%d\n",
			pos, verInst, recType, recLen, ver, inst)

		if recType >= 0xF000 && recType <= 0xF200 {
			// Valid OfficeArt record
			if ver == 0x0F {
				// Container - enter it
				pos += 8
			} else {
				pos += 8 + int(recLen)
			}
		} else {
			fmt.Printf("  Not a valid OfficeArt record type, stopping\n")
			break
		}

		if pos > 200 {
			fmt.Println("  ... (truncated)")
			break
		}
	}

	// The DggInfo in Word documents is actually a Drawing Group Container
	// that wraps the OfficeArt data. In Word, it starts with a DggContainer (0xF000)
	// Let me check if the first record is actually a DggContainer
	if len(dggData) >= 8 {
		verInst := binary.LittleEndian.Uint16(dggData[0:])
		recType := binary.LittleEndian.Uint16(dggData[2:])
		recLen := binary.LittleEndian.Uint32(dggData[4:])
		ver := verInst & 0x0F
		fmt.Printf("\nFirst record: type=0x%04X ver=%d len=%d\n", recType, ver, recLen)
		
		if recType == 0xF000 && ver == 0x0F {
			fmt.Println("It IS a DggContainer!")
			// Parse inside
			innerEnd := 8 + int(recLen)
			if innerEnd > len(dggData) {
				innerEnd = len(dggData)
			}
			parseInner(dggData, 8, innerEnd, 1)
		}
	}
}

func parseInner(data []byte, offset, end, depth int) {
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

		childEnd := offset + 8 + int(recLen)
		if childEnd > end {
			childEnd = end
		}

		if ver == 0x0F {
			parseInner(data, offset+8, childEnd, depth+1)
		} else if recType == 0xF00A && int(recLen) >= 8 {
			spid := binary.LittleEndian.Uint32(data[offset+8:])
			flags := binary.LittleEndian.Uint32(data[offset+12:])
			fmt.Printf("%s  spid=%d flags=0x%08X\n", indent, spid, flags)
		} else if recType == 0xF00B {
			// OPT - parse properties
			propOff := offset + 8
			for p := uint16(0); p < inst && propOff+6 <= childEnd; p++ {
				propID := binary.LittleEndian.Uint16(data[propOff:])
				propVal := binary.LittleEndian.Uint32(data[propOff+2:])
				pid := propID & 0x3FFF
				switch pid {
				case 0x0080:
					fmt.Printf("%s  lTxid=%d\n", indent, propVal)
				case 0x0104:
					fmt.Printf("%s  pib=%d\n", indent, propVal)
				}
				propOff += 6
			}
		} else if recType == 0xF006 && int(recLen) >= 16 {
			// Dgg record
			spidMax := binary.LittleEndian.Uint32(data[offset+8:])
			cidcl := binary.LittleEndian.Uint32(data[offset+12:])
			cspSaved := binary.LittleEndian.Uint32(data[offset+16:])
			cdgSaved := binary.LittleEndian.Uint32(data[offset+20:])
			fmt.Printf("%s  spidMax=%d cidcl=%d cspSaved=%d cdgSaved=%d\n",
				indent, spidMax, cidcl, cspSaved, cdgSaved)
		}

		offset = childEnd
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
