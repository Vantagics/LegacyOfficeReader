package main

import (
	"encoding/binary"
	"fmt"
	"os"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	stat, _ := f.Stat()
	data := make([]byte, stat.Size())
	f.Read(data)

	sectorShift := binary.LittleEndian.Uint16(data[30:])
	sectorSize := uint32(1) << sectorShift
	firstDirSector := binary.LittleEndian.Uint32(data[48:])

	var fatSectors []uint32
	for i := 0; i < 109; i++ {
		val := binary.LittleEndian.Uint32(data[76+i*4:])
		if val == 0xFFFFFFFE || val == 0xFFFFFFFF {
			break
		}
		fatSectors = append(fatSectors, val)
	}
	fat := make(map[uint32]uint32)
	idx := uint32(0)
	for _, fs := range fatSectors {
		fatOff := (fs + 1) * sectorSize
		for i := uint32(0); i < sectorSize/4; i++ {
			off := fatOff + i*4
			if int(off+4) > len(data) {
				break
			}
			fat[idx] = binary.LittleEndian.Uint32(data[off:])
			idx++
		}
	}

	type dirEntry struct {
		name        string
		startSector uint32
		size        uint32
	}
	var dirs []dirEntry
	dirSector := firstDirSector
	for dirSector != 0xFFFFFFFE && dirSector != 0xFFFFFFFF {
		dirOff := (dirSector + 1) * sectorSize
		for i := uint32(0); i < sectorSize/128; i++ {
			entryOff := dirOff + i*128
			if int(entryOff+128) > len(data) {
				break
			}
			nameSize := binary.LittleEndian.Uint16(data[entryOff+64:])
			name := ""
			for j := uint16(0); j+2 <= nameSize; j += 2 {
				ch := binary.LittleEndian.Uint16(data[entryOff+uint32(j):])
				if ch == 0 {
					break
				}
				name += string(rune(ch))
			}
			objType := data[entryOff+66]
			startSec := binary.LittleEndian.Uint32(data[entryOff+116:])
			size := binary.LittleEndian.Uint32(data[entryOff+120:])
			if objType != 0 {
				dirs = append(dirs, dirEntry{name: name, startSector: startSec, size: size})
			}
		}
		dirSector = fat[dirSector]
	}

	var wdStart, wdSize, tableStart, tableSize uint32
	for _, d := range dirs {
		if d.name == "WordDocument" {
			wdStart = d.startSector
			wdSize = d.size
		}
		if d.name == "1Table" {
			tableStart = d.startSector
			tableSize = d.size
		}
	}

	wdData := readChain(data, fat, wdStart, wdSize, sectorSize)
	tableData := readChain(data, fat, tableStart, tableSize, sectorSize)

	// Check all FIB offsets related to OfficeArt
	// FibRgFcLcb97 starts at offset 0x009A in FIB
	// DggInfo is at FibRgFcLcb97 offset 0x150 (relative to 0x009A)
	// That's absolute 0x01EA in the FIB
	
	// But wait - the FIB version matters. Let's check the FIB structure more carefully.
	fibVer := binary.LittleEndian.Uint16(wdData[0x0002:])
	fmt.Printf("FIB version: 0x%04X\n", fibVer)
	
	// FibBase is 32 bytes (0x0000-0x001F)
	// csw at 0x0020 (count of shorts in FibRgW)
	csw := binary.LittleEndian.Uint16(wdData[0x0020:])
	fmt.Printf("csw (FibRgW count): %d\n", csw)
	
	// FibRgW starts at 0x0022, size = csw*2 bytes
	fibRgWEnd := 0x0022 + int(csw)*2
	
	// cslw at fibRgWEnd (count of longs in FibRgLw)
	cslw := binary.LittleEndian.Uint16(wdData[fibRgWEnd:])
	fmt.Printf("cslw (FibRgLw count): %d\n", cslw)
	
	// FibRgLw starts at fibRgWEnd+2
	fibRgLwStart := fibRgWEnd + 2
	fibRgLwEnd := fibRgLwStart + int(cslw)*4
	
	// ccpText is at FibRgLw offset 3 (0-indexed)
	fmt.Printf("FibRgLw starts at 0x%04X, ends at 0x%04X\n", fibRgLwStart, fibRgLwEnd)
	
	// cbRgFcLcb at fibRgLwEnd
	cbRgFcLcb := binary.LittleEndian.Uint16(wdData[fibRgLwEnd:])
	fmt.Printf("cbRgFcLcb (FibRgFcLcb count): %d\n", cbRgFcLcb)
	
	// FibRgFcLcb starts at fibRgLwEnd+2
	fibRgFcLcbStart := fibRgLwEnd + 2
	fmt.Printf("FibRgFcLcb starts at 0x%04X\n", fibRgFcLcbStart)
	
	// DggInfo is at FibRgFcLcb97 index 168/169 (fc/lcb pair)
	// Each pair is 8 bytes (fc=4, lcb=4)
	// Index 168 = offset 168*4 = 672 from start of FibRgFcLcb
	// But actually, the pairs are sequential: fc0, lcb0, fc1, lcb1, ...
	// So pair N is at offset N*8
	
	// Let me just check the known offsets
	// PlcfHdd: pair 31 (fc at offset 31*8=248, lcb at 252)
	plcfHddFC := binary.LittleEndian.Uint32(wdData[fibRgFcLcbStart+248:])
	plcfHddLCB := binary.LittleEndian.Uint32(wdData[fibRgFcLcbStart+252:])
	fmt.Printf("PlcfHdd: fc=%d lcb=%d (at offset 0x%04X)\n", plcfHddFC, plcfHddLCB, fibRgFcLcbStart+248)
	
	// PlcSpaMom: pair 88
	plcSpaMomFC := binary.LittleEndian.Uint32(wdData[fibRgFcLcbStart+88*8:])
	plcSpaMomLCB := binary.LittleEndian.Uint32(wdData[fibRgFcLcbStart+88*8+4:])
	fmt.Printf("PlcSpaMom: fc=%d lcb=%d\n", plcSpaMomFC, plcSpaMomLCB)
	
	// PlcSpaHdr: pair 89
	plcSpaHdrFC := binary.LittleEndian.Uint32(wdData[fibRgFcLcbStart+89*8:])
	plcSpaHdrLCB := binary.LittleEndian.Uint32(wdData[fibRgFcLcbStart+89*8+4:])
	fmt.Printf("PlcSpaHdr: fc=%d lcb=%d\n", plcSpaHdrFC, plcSpaHdrLCB)
	
	// fcDggInfo: pair 90
	dggInfoFC := binary.LittleEndian.Uint32(wdData[fibRgFcLcbStart+90*8:])
	dggInfoLCB := binary.LittleEndian.Uint32(wdData[fibRgFcLcbStart+90*8+4:])
	fmt.Printf("DggInfo: fc=%d lcb=%d\n", dggInfoFC, dggInfoLCB)
	
	// The DggInfo is in the Table stream
	if dggInfoLCB > 0 && int(dggInfoFC+dggInfoLCB) <= len(tableData) {
		dggData := tableData[dggInfoFC : dggInfoFC+dggInfoLCB]
		fmt.Printf("DggInfo data available: %d bytes\n", len(dggData))
		// First 8 bytes
		if len(dggData) >= 8 {
			vi := binary.LittleEndian.Uint16(dggData[0:])
			rt := binary.LittleEndian.Uint16(dggData[2:])
			rl := binary.LittleEndian.Uint32(dggData[4:])
			fmt.Printf("  First record: ver/inst=0x%04X type=0x%04X len=%d\n", vi, rt, rl)
		}
	} else {
		fmt.Println("DggInfo not available or empty")
	}
	
	// Check if there's OfficeArt in the WordDocument stream itself
	// The OfficeArt Drawing Group Container is typically in the Table stream
	// But header shapes might reference images via BSE indices
	
	// Let's parse PlcSpaHdr to get the SPIDs
	if plcSpaHdrLCB > 0 {
		spaData := tableData[plcSpaHdrFC : plcSpaHdrFC+plcSpaHdrLCB]
		n := (plcSpaHdrLCB - 4) / 30
		fmt.Printf("\nPlcSpaHdr: %d shapes\n", n)
		for i := uint32(0); i < n; i++ {
			cp := binary.LittleEndian.Uint32(spaData[i*4:])
			spaOff := (n+1)*4 + i*26
			if spaOff+26 > uint32(len(spaData)) {
				break
			}
			spid := binary.LittleEndian.Uint32(spaData[spaOff:])
			xaLeft := int32(binary.LittleEndian.Uint32(spaData[spaOff+4:]))
			yaTop := int32(binary.LittleEndian.Uint32(spaData[spaOff+8:]))
			xaRight := int32(binary.LittleEndian.Uint32(spaData[spaOff+12:]))
			yaBottom := int32(binary.LittleEndian.Uint32(spaData[spaOff+16:]))
			flags := binary.LittleEndian.Uint16(spaData[spaOff+20:])
			fmt.Printf("  Shape[%d]: cp=%d spid=%d rect=(%d,%d,%d,%d) flags=0x%04X\n",
				i, cp, spid, xaLeft, yaTop, xaRight, yaBottom, flags)
		}
	}
	
	// Now scan the DggInfo for these SPIDs
	// The OfficeArt Drawing Group Container contains:
	// 1. DggContainer (0xF000) - contains DGG, BStoreContainer, etc.
	// 2. DgContainer(s) (0xF002) - contains drawing groups with shapes
	
	// For header shapes, they're in a separate DgContainer
	// Let's scan the entire DggInfo
	if dggInfoLCB > 0 && int(dggInfoFC+dggInfoLCB) <= len(tableData) {
		dggData := tableData[dggInfoFC : dggInfoFC+dggInfoLCB]
		fmt.Println("\n=== Full OfficeArt scan ===")
		scanAllShapes(dggData, 0, len(dggData), 0)
	}
}

func scanAllShapes(data []byte, offset, end, depth int) {
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

		if ver == 0x0F {
			// Container
			if recType == 0xF004 {
				// SpContainer
				spid, props := parseSpProps(data, offset+8, childEnd, inst)
				if spid != 0 {
					fmt.Printf("  SPID=%d props: %v\n", spid, props)
				}
			}
			scanAllShapes(data, offset+8, childEnd, depth+1)
		}

		offset = childEnd
	}
}

func parseSpProps(data []byte, offset, end int, parentInst uint16) (uint32, map[string]uint32) {
	var spid uint32
	props := make(map[string]uint32)
	
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
			flags := binary.LittleEndian.Uint32(recData[4:])
			props["flags"] = flags
		}

		if recType == 0xF00B || recType == 0xF122 || recType == 0xF11E {
			for p := uint16(0); p < inst; p++ {
				off := int(p) * 6
				if off+6 > len(recData) {
					break
				}
				propID := binary.LittleEndian.Uint16(recData[off:])
				propVal := binary.LittleEndian.Uint32(recData[off+2:])
				pid := propID & 0x3FFF
				
				switch pid {
				case 0x0004: // pib
					props["pib"] = propVal
				case 0x0080: // lTxid
					props["lTxid"] = propVal
				case 0x0104: // pib (fill)
					props["fillPib"] = propVal
				case 0x0186: // fillBlip
					props["fillBlip"] = propVal
				case 0x0181: // fillColor
					props["fillColor"] = propVal
				case 0x0180: // fillType
					props["fillType"] = propVal
				}
			}
		}

		offset = childEnd
	}
	return spid, props
}

func readChain(data []byte, fat map[uint32]uint32, startSector, size, sectorSize uint32) []byte {
	var result []byte
	sector := startSector
	remaining := size
	for remaining > 0 && sector != 0xFFFFFFFE && sector != 0xFFFFFFFF {
		off := (sector + 1) * sectorSize
		readSize := sectorSize
		if readSize > remaining {
			readSize = remaining
		}
		if int(off+readSize) > len(data) {
			break
		}
		result = append(result, data[off:off+readSize]...)
		remaining -= readSize
		sector = fat[sector]
	}
	return result
}
