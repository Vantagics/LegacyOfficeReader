package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
)

func main() {
	f, _ := os.Open("testfie/test.doc")
	defer f.Close()

	adaptor, _ := cfb.OpenReader(f)
	var wordDoc, table1, root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "WordDocument":
			wordDoc = dir
		case "1Table":
			table1 = dir
		case "Root Entry":
			root = dir
		}
	}

	wdReader, _ := adaptor.OpenObject(wordDoc, root)
	wdSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wdData := make([]byte, wdSize)
	wdReader.Read(wdData)

	tReader, _ := adaptor.OpenObject(table1, root)
	tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tData := make([]byte, tSize)
	tReader.Read(tData)

	// Navigate FIB like fib.go does
	offset := 0x20
	csw := binary.LittleEndian.Uint16(wdData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wdData[offset:])
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wdData[offset:])
	offset += 2

	fmt.Printf("FibRgFcLcb: %d uint32s, starts at offset 0x%X\n", cbRgFcLcb, offset)

	readFcLcb := func(idx int) uint32 {
		if int(cbRgFcLcb) <= idx {
			return 0
		}
		off := offset + idx*4
		return binary.LittleEndian.Uint32(wdData[off:])
	}

	fcPlcSpaMom := readFcLcb(80)
	lcbPlcSpaMom := readFcLcb(81)
	fcDggInfo := readFcLcb(100)
	lcbDggInfo := readFcLcb(101)

	fmt.Printf("fcPlcSpaMom=%d lcbPlcSpaMom=%d\n", fcPlcSpaMom, lcbPlcSpaMom)
	fmt.Printf("fcDggInfo=%d lcbDggInfo=%d\n", fcDggInfo, lcbDggInfo)

	// Parse PlcSpaMom
	if lcbPlcSpaMom > 0 && uint64(fcPlcSpaMom)+uint64(lcbPlcSpaMom) <= uint64(len(tData)) {
		spaData := tData[fcPlcSpaMom : fcPlcSpaMom+lcbPlcSpaMom]
		n := (lcbPlcSpaMom - 4) / 30
		fmt.Printf("\nPlcSpaMom: %d SPAs\n", n)
		for i := uint32(0); i < n; i++ {
			cp := binary.LittleEndian.Uint32(spaData[i*4:])
			spaOff := (n+1)*4 + i*26
			spid := binary.LittleEndian.Uint32(spaData[spaOff:])
			fmt.Printf("  SPA[%d]: CP=%d SPID=%d\n", i, cp, spid)
		}
	}

	// Parse DggInfo shapes
	if lcbDggInfo > 0 && uint64(fcDggInfo)+uint64(lcbDggInfo) <= uint64(len(tData)) {
		data := tData[fcDggInfo : fcDggInfo+lcbDggInfo]
		fmt.Printf("\nDggInfo: %d bytes\n", len(data))
		if len(data) >= 8 {
			vi := binary.LittleEndian.Uint16(data[0:])
			rt := binary.LittleEndian.Uint16(data[2:])
			rl := binary.LittleEndian.Uint32(data[4:])
			fmt.Printf("  First record: ver=0x%X type=0x%04X len=%d\n", vi&0x0F, rt, rl)

			pos := 8 + int(rl) // skip DggContainer

			dgIdx := 0
			for pos < len(data) {
				found := false
				for pos+8 <= len(data) {
					vi2 := binary.LittleEndian.Uint16(data[pos:])
					rt2 := binary.LittleEndian.Uint16(data[pos+2:])
					v2 := vi2 & 0x0F
					if rt2 == 0xF002 && v2 == 0x0F {
						found = true
						break
					}
					pos++
				}
				if !found {
					break
				}

				rl2 := binary.LittleEndian.Uint32(data[pos+4:])
				containerEnd := pos + 8 + int(rl2)
				if containerEnd > len(data) {
					containerEnd = len(data)
				}

				fmt.Printf("\n  DgContainer #%d (offset %d, len %d):\n", dgIdx, pos, rl2)
				printSpContainers(data, pos+8, containerEnd, "    ")

				pos = containerEnd
				dgIdx++
			}
		}
	} else {
		fmt.Println("\nNo DggInfo data")
	}
}

func printSpContainers(data []byte, offset, end int, indent string) {
	for offset+8 <= end {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])
		ver := verInst & 0x0F

		childEnd := offset + 8 + int(recLen)
		if childEnd > end {
			childEnd = end
		}

		if ver == 0x0F {
			if recType == 0xF004 {
				spid, pib, shapeType, txid := parseSpDetail(data, offset+8, childEnd)
				extra := ""
				if txid != 0 {
					extra = fmt.Sprintf(" txid=%d", txid)
				}
				fmt.Printf("%sSpContainer: SPID=%d pib=%d shapeType=0x%04X%s\n", indent, spid, pib, shapeType, extra)
			} else {
				printSpContainers(data, offset+8, childEnd, indent)
			}
		}

		offset = childEnd
	}
}

func parseSpDetail(data []byte, offset, end int) (spid, pib uint32, shapeType uint16, txid uint32) {
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
			shapeType = inst
		}
		if recType == 0xF00B {
			for p := uint16(0); p < inst; p++ {
				off := int(p) * 6
				if off+6 > len(recData) {
					break
				}
				propID := binary.LittleEndian.Uint16(recData[off:])
				propVal := binary.LittleEndian.Uint32(recData[off+2:])
				pid := propID & 0x3FFF
				if pid == 260 {
					pib = propVal
				}
				if pid == 0x0080 { // lTxid
					txid = propVal
				}
			}
		}
		offset = childEnd
	}
	return
}
