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

	dggData := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]

	// The OfficeArtContent structure is:
	// 1. DggContainer (record type 0xF000)
	// 2. DgContainer for main document (record type 0xF002)
	// 3. DgContainer for header document (record type 0xF002) - optional
	// These are sequential top-level records

	pos := 0
	recordNum := 0
	for pos+8 <= len(dggData) {
		verInst := binary.LittleEndian.Uint16(dggData[pos:])
		recType := binary.LittleEndian.Uint16(dggData[pos+2:])
		recLen := binary.LittleEndian.Uint32(dggData[pos+4:])
		ver := verInst & 0x0F

		name := msoRecordName(recType)
		fmt.Printf("Record[%d] at offset %d: type=0x%04X (%s) ver=%d len=%d\n",
			recordNum, pos, recType, name, ver, recLen)

		if recType == 0xF002 { // DgContainer
			fmt.Printf("  Parsing DgContainer...\n")
			innerEnd := pos + 8 + int(recLen)
			if innerEnd > len(dggData) {
				innerEnd = len(dggData)
			}
			parseDgContainer(dggData, pos+8, innerEnd)
		}

		if ver == 0x0F {
			pos += 8 + int(recLen)
		} else {
			pos += 8 + int(recLen)
		}
		recordNum++
		if recordNum > 10 {
			break
		}
	}
}

func parseDgContainer(data []byte, offset, end int) {
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

		if recType == 0xF004 { // SpContainer
			fmt.Printf("    SpContainer at %d len=%d\n", offset, recLen)
			parseSpContainer(data, offset+8, childEnd)
		} else if recType == 0xF003 { // SpgrContainer
			fmt.Printf("    SpgrContainer at %d len=%d\n", offset, recLen)
			// Parse inside for SpContainers
			parseDgContainer(data, offset+8, childEnd)
		} else if recType == 0xF008 && int(recLen) >= 8 { // Dg
			csp := binary.LittleEndian.Uint32(data[offset+8:])
			spidCur := binary.LittleEndian.Uint32(data[offset+12:])
			fmt.Printf("    [Dg] csp=%d spidCur=%d\n", csp, spidCur)
		} else {
			fmt.Printf("    [0x%04X] %s ver=%d inst=%d len=%d\n", recType, name, ver, inst, recLen)
		}

		if ver == 0x0F {
			offset = childEnd
		} else {
			offset = childEnd
		}
	}
}

func parseSpContainer(data []byte, offset, end int) {
	var spid uint32
	var pib uint32
	var txid uint32
	var hasClientTextbox bool

	for offset+8 <= end {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])
		ver := verInst & 0x0F
		inst := verInst >> 4

		childEnd := offset + 8 + int(recLen)
		if childEnd > end {
			childEnd = end
		}
		recData := data[offset+8 : childEnd]

		if recType == 0xF00A && len(recData) >= 8 { // Sp
			spid = binary.LittleEndian.Uint32(recData[0:])
		}
		if recType == 0xF00D { // ClientTextbox
			hasClientTextbox = true
		}
		if recType == 0xF00B { // OPT
			for p := uint16(0); p < inst; p++ {
				off := int(p) * 6
				if off+6 > len(recData) {
					break
				}
				propID := binary.LittleEndian.Uint16(recData[off:])
				propVal := binary.LittleEndian.Uint32(recData[off+2:])
				pid := propID & 0x3FFF
				if pid == 0x0104 { pib = propVal }
				if pid == 0x0080 { txid = propVal }
			}
		}

		if ver == 0x0F {
			// Recurse
			s2, p2, t2, ct2 := parseSpContainerInner(data, offset+8, childEnd)
			if s2 != 0 { spid = s2 }
			if p2 != 0 { pib = p2 }
			if t2 != 0 { txid = t2 }
			if ct2 { hasClientTextbox = true }
		}

		offset = childEnd
	}

	fmt.Printf("      spid=%d pib=%d txid=%d clientTextbox=%v\n", spid, pib, txid, hasClientTextbox)
}

func parseSpContainerInner(data []byte, offset, end int) (spid, pib, txid uint32, hasClientTextbox bool) {
	for offset+8 <= end {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])
		inst := verInst >> 4

		childEnd := offset + 8 + int(recLen)
		if childEnd > end {
			childEnd = end
		}
		recData := data[offset+8 : childEnd]

		if recType == 0xF00A && len(recData) >= 8 {
			spid = binary.LittleEndian.Uint32(recData[0:])
		}
		if recType == 0xF00D {
			hasClientTextbox = true
		}
		if recType == 0xF00B {
			for p := uint16(0); p < inst; p++ {
				off := int(p) * 6
				if off+6 > len(recData) { break }
				propID := binary.LittleEndian.Uint16(recData[off:])
				propVal := binary.LittleEndian.Uint32(recData[off+2:])
				pid := propID & 0x3FFF
				if pid == 0x0104 { pib = propVal }
				if pid == 0x0080 { txid = propVal }
			}
		}
		offset = childEnd
	}
	return
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
	default: return fmt.Sprintf("Unknown(0x%04X)", t)
	}
}
