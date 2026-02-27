package main

import (
	"encoding/binary"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	// Read the whole file
	stat, _ := f.Stat()
	data := make([]byte, stat.Size())
	f.Read(data)

	// Parse CFB to get streams
	sectorShift := binary.LittleEndian.Uint16(data[30:])
	sectorSize := uint32(1) << sectorShift
	firstDirSector := binary.LittleEndian.Uint32(data[48:])

	// Build FAT
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

	// Read directory entries
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

	// Get DggInfo (OfficeArt data)
	fcDggInfo := binary.LittleEndian.Uint32(wdData[0x01EA:])
	lcbDggInfo := binary.LittleEndian.Uint32(wdData[0x01EE:])
	fmt.Printf("DggInfo: fc=%d lcb=%d\n", fcDggInfo, lcbDggInfo)

	if lcbDggInfo == 0 {
		fmt.Println("No DggInfo")
		return
	}

	dggData := tableData[fcDggInfo : fcDggInfo+lcbDggInfo]

	// Parse OfficeArt to find shapes with SPIDs 2049, 2050, 2051
	// and determine their BSE image references
	targetSPIDs := map[uint32]bool{2049: true, 2050: true, 2051: true}

	fmt.Println("\n=== Scanning OfficeArt for header shapes ===")
	scanOfficeArt(dggData, 0, len(dggData), 0, targetSPIDs)

	// Also check the BSE table to see which images are available
	fmt.Println("\n=== BSE Table ===")
	scanBSE(dggData)

	// Now use the doc package to verify
	fmt.Println("\n=== Using doc package ===")
	f.Seek(0, 0)
	d, err := doc.OpenReader(f)
	if err != nil {
		panic(err)
	}
	images := d.GetImages()
	for i, img := range images {
		fmt.Printf("  BSE[%d]: format=%d size=%d\n", i, img.Format, len(img.Data))
	}
}

func scanOfficeArt(data []byte, offset, end, depth int, targets map[uint32]bool) {
	indent := strings.Repeat("  ", depth)
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

		typeName := fmt.Sprintf("0x%04X", recType)
		switch recType {
		case 0xF000:
			typeName = "DggContainer"
		case 0xF001:
			typeName = "BStoreContainer"
		case 0xF002:
			typeName = "DgContainer"
		case 0xF003:
			typeName = "SpgrContainer"
		case 0xF004:
			typeName = "SpContainer"
		case 0xF006:
			typeName = "FDGGBlock"
		case 0xF007:
			typeName = "BSE"
		case 0xF008:
			typeName = "FDG"
		case 0xF009:
			typeName = "FSPGR"
		case 0xF00A:
			typeName = "FSP"
		case 0xF00B:
			typeName = "FOPT"
		case 0xF00D:
			typeName = "ClientAnchor"
		case 0xF010:
			typeName = "ClientData"
		case 0xF011:
			typeName = "ClientTextbox"
		case 0xF01E:
			typeName = "SplitMenuColors"
		case 0xF11E:
			typeName = "TertiaryFOPT"
		case 0xF122:
			typeName = "SecondaryFOPT"
		}

		_ = typeName
		if ver == 0x0F {
			// Container - recurse
			if recType == 0xF004 {
				// SpContainer - parse for SPID and properties
				spid, bseIdx, hasBse := parseSpContainer(data, offset+8, childEnd, inst)
				if targets[spid] {
					fmt.Printf("%sSPID=%d bseIdx=%d hasBse=%v\n", indent, spid, bseIdx, hasBse)
				}
			}
			scanOfficeArt(data, offset+8, childEnd, depth+1, targets)
		}

		offset = childEnd
	}
}

func parseSpContainer(data []byte, offset, end int, parentInst uint16) (spid uint32, bseIdx int, hasBse bool) {
	bseIdx = -1
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
			// FSP record
			spid = binary.LittleEndian.Uint32(recData[0:])
		}

		if recType == 0xF00B || recType == 0xF122 || recType == 0xF11E {
			// FOPT/SecondaryFOPT/TertiaryFOPT - parse properties
			for p := uint16(0); p < inst; p++ {
				off := int(p) * 6
				if off+6 > len(recData) {
					break
				}
				propID := binary.LittleEndian.Uint16(recData[off:])
				propVal := binary.LittleEndian.Uint32(recData[off+2:])
				pid := propID & 0x3FFF

				// pib (picture index) properties
				// 0x0104 = pib (blip index, 1-based)
				// 0x0105 = pibName
				// 0x0180 = fillType
				// 0x0181 = fillColor
				// 0x0186 = fillBlip (fill pattern blip, 1-based BSE index)
				if pid == 0x0104 {
					bseIdx = int(propVal) - 1 // convert to 0-based
					hasBse = true
					fmt.Printf("    SPID=%d: pib (0x0104) = %d (BSE index %d)\n", spid, propVal, bseIdx)
				}
				if pid == 0x0186 {
					fmt.Printf("    SPID=%d: fillBlip (0x0186) = %d\n", spid, propVal)
				}
				if pid == 0x0080 {
					fmt.Printf("    SPID=%d: lTxid (0x0080) = %d\n", spid, propVal)
				}
			}
		}

		offset = childEnd
	}
	return
}

func scanBSE(data []byte) {
	// Find BStoreContainer (0xF001)
	offset := 0
	for offset+8 <= len(data) {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])
		ver := verInst & 0x0F

		childEnd := offset + 8 + int(recLen)
		if childEnd > len(data) {
			childEnd = len(data)
		}

		if recType == 0xF001 && ver == 0x0F {
			// BStoreContainer - enumerate BSE entries
			bseOffset := offset + 8
			bseIdx := 0
			for bseOffset+8 <= childEnd {
				bseVI := binary.LittleEndian.Uint16(data[bseOffset:])
				bseRT := binary.LittleEndian.Uint16(data[bseOffset+2:])
				bseRL := binary.LittleEndian.Uint32(data[bseOffset+4:])
				bseInst := bseVI >> 4
				_ = bseRT

				bseEnd := bseOffset + 8 + int(bseRL)
				if bseEnd > childEnd {
					bseEnd = childEnd
				}

				// BSE record: first 36 bytes are FBSE header
				if bseOffset+8+36 <= bseEnd {
					bseData := data[bseOffset+8:]
					btWin32 := bseData[0]
					_ = bseData[1] // btMacOS
					cRef := binary.LittleEndian.Uint32(bseData[24:])
					
					formatNames := map[byte]string{
						0: "ERROR", 1: "UNKNOWN", 2: "EMF", 3: "WMF",
						4: "PICT", 5: "JPEG", 6: "PNG", 7: "DIB",
						8: "TIFF", 9: "CMYK_JPEG",
					}
					fname := formatNames[btWin32]
					if fname == "" {
						fname = fmt.Sprintf("type_%d", btWin32)
					}

					fmt.Printf("  BSE[%d]: inst=0x%04X format=%s(%d) cRef=%d\n",
						bseIdx, bseInst, fname, btWin32, cRef)
				}

				bseOffset = bseEnd
				bseIdx++
			}
			return
		}

		if ver == 0x0F {
			// Container - recurse into it
			offset += 8
		} else {
			offset = childEnd
		}
	}
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
