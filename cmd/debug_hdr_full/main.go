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

	d, err := doc.OpenReader(f)
	if err != nil {
		panic(err)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	fmt.Printf("Headers: %d, Footers: %d\n", len(fc.Headers), len(fc.Footers))
	for i, h := range fc.Headers {
		fmt.Printf("Header[%d]: %q\n", i, h)
	}
	for i, h := range fc.HeadersRaw {
		fmt.Printf("HeaderRaw[%d]: %q\n", i, h)
	}
	for i, f := range fc.Footers {
		fmt.Printf("Footer[%d]: %q\n", i, f)
	}
	for i, f := range fc.FootersRaw {
		fmt.Printf("FooterRaw[%d]: %q\n", i, f)
	}

	// Now let's directly parse the binary to see ALL header/footer stories
	// including ones with only images (0x01/0x08 chars)
	fmt.Println("\n=== Direct Binary Analysis ===")
	
	// Re-open and read raw
	f.Seek(0, 0)
	rawData := make([]byte, 100*1024*1024)
	n, _ := f.Read(rawData)
	rawData = rawData[:n]
	_ = rawData

	// Use the doc package's internal data - we need to access the raw streams
	// Let's use the exported methods
	text := d.GetText()
	fmt.Printf("Full text length: %d chars\n", len(text))
	
	// Check for images
	images := d.GetImages()
	fmt.Printf("Total images: %d\n", len(images))
	for i, img := range images {
		fmt.Printf("  Image[%d]: format=%d size=%d bytes\n", i, img.Format, len(img.Data))
	}

	// Let's look at the raw binary to understand header structure
	// We need to parse the DOC file at a lower level
	fmt.Println("\n=== Low-level header/footer analysis ===")
	f.Seek(0, 0)
	analyzeLowLevel(f)
}

func analyzeLowLevel(f *os.File) {
	// Read WordDocument stream header
	f.Seek(0, 0)
	
	// Read the CFB and find WordDocument
	// Use cfb package
	f.Seek(0, 0)
	
	// Read the whole file
	f.Seek(0, 0)
	stat, _ := f.Stat()
	data := make([]byte, stat.Size())
	f.Read(data)
	
	// CFB header: first 512 bytes
	// Sector size: 2^(data[30:32])
	sectorShift := binary.LittleEndian.Uint16(data[30:])
	sectorSize := uint32(1) << sectorShift
	fmt.Printf("Sector size: %d\n", sectorSize)
	
	// First directory sector location: data[48:52]
	firstDirSector := binary.LittleEndian.Uint32(data[48:])
	fmt.Printf("First dir sector: %d\n", firstDirSector)
	
	// Read FAT to follow chains
	// DIFAT entries start at offset 76 in header, 109 entries
	var fatSectors []uint32
	for i := 0; i < 109; i++ {
		off := 76 + i*4
		val := binary.LittleEndian.Uint32(data[off:])
		if val == 0xFFFFFFFE || val == 0xFFFFFFFF {
			break
		}
		fatSectors = append(fatSectors, val)
	}
	
	// Build FAT table
	fat := make(map[uint32]uint32)
	for _, fs := range fatSectors {
		fatOff := (fs + 1) * sectorSize
		for i := uint32(0); i < sectorSize/4; i++ {
			off := fatOff + i*4
			if int(off+4) > len(data) {
				break
			}
			fat[uint32(len(fat))] = binary.LittleEndian.Uint32(data[off:])
		}
	}
	
	// Read directory entries
	type dirEntry struct {
		name string
		startSector uint32
		size uint32
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
			// Read name (UTF-16LE)
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
	
	fmt.Println("Directory entries:")
	for _, d := range dirs {
		fmt.Printf("  %s: start=%d size=%d\n", d.name, d.startSector, d.size)
	}
	
	// Find WordDocument and 1Table
	var wdStart, wdSize uint32
	var tableStart, tableSize uint32
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
	
	// Read WordDocument stream
	wdData := readFatChain(data, fat, wdStart, wdSize, sectorSize)
	fmt.Printf("WordDocument: %d bytes\n", len(wdData))
	
	// Read 1Table stream
	tableData := readFatChain(data, fat, tableStart, tableSize, sectorSize)
	fmt.Printf("1Table: %d bytes\n", len(tableData))
	
	// Parse FIB
	ccpText := binary.LittleEndian.Uint32(wdData[0x004C:])
	ccpFtn := binary.LittleEndian.Uint32(wdData[0x0050:])
	ccpHdd := binary.LittleEndian.Uint32(wdData[0x0054:])
	ccpAtn := binary.LittleEndian.Uint32(wdData[0x0058:])
	ccpEdn := binary.LittleEndian.Uint32(wdData[0x005C:])
	ccpTxbx := binary.LittleEndian.Uint32(wdData[0x0060:])
	ccpHdrTxbx := binary.LittleEndian.Uint32(wdData[0x0064:])
	
	fmt.Printf("\nFIB: ccpText=%d ccpFtn=%d ccpHdd=%d ccpAtn=%d ccpEdn=%d ccpTxbx=%d ccpHdrTxbx=%d\n",
		ccpText, ccpFtn, ccpHdd, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx)
	
	// PlcfHdd
	fcPlcfHdd := binary.LittleEndian.Uint32(wdData[0x00F2:])
	lcbPlcfHdd := binary.LittleEndian.Uint32(wdData[0x00F6:])
	fmt.Printf("PlcfHdd: fc=%d lcb=%d\n", fcPlcfHdd, lcbPlcfHdd)
	
	// PlcSpaHdr (shapes in headers/footers)
	fcPlcSpaHdr := binary.LittleEndian.Uint32(wdData[0x01E2:])
	lcbPlcSpaHdr := binary.LittleEndian.Uint32(wdData[0x01E6:])
	fmt.Printf("PlcSpaHdr: fc=%d lcb=%d\n", fcPlcSpaHdr, lcbPlcSpaHdr)
	
	// Extract full text via piece table
	fcClx := binary.LittleEndian.Uint32(wdData[0x01A2:])
	lcbClx := binary.LittleEndian.Uint32(wdData[0x01A6:])
	fullText := extractPieceText(wdData, tableData, fcClx, lcbClx)
	fullRunes := []rune(fullText)
	fmt.Printf("Full text: %d runes\n", len(fullRunes))
	
	// Parse PlcfHdd
	if lcbPlcfHdd > 0 {
		plcData := tableData[fcPlcfHdd : fcPlcfHdd+lcbPlcfHdd]
		nCPs := lcbPlcfHdd / 4
		cps := make([]uint32, nCPs)
		for i := uint32(0); i < nCPs; i++ {
			cps[i] = binary.LittleEndian.Uint32(plcData[i*4:])
		}
		
		hddStart := ccpText + ccpFtn
		hddEnd := hddStart + ccpHdd
		
		storyNames := []string{
			"even-hdr", "odd-hdr", "even-ftr", "odd-ftr", "first-hdr", "first-ftr",
		}
		
		fmt.Printf("\nPlcfHdd: %d CPs, hddStart=%d, hddEnd=%d\n", nCPs, hddStart, hddEnd)
		for i := 0; i+1 < int(nCPs); i++ {
			cpStart := hddStart + cps[i]
			cpEnd := hddStart + cps[i+1]
			if cpEnd > hddEnd {
				cpEnd = hddEnd
			}
			
			sectionIdx := i / 6
			storyIdx := i % 6
			name := storyNames[storyIdx]
			
			storyLen := int(cpEnd) - int(cpStart)
			if cpStart >= uint32(len(fullRunes)) || cpEnd > uint32(len(fullRunes)) || cpStart >= cpEnd {
				fmt.Printf("  [%d] sec=%d %s: EMPTY (cp %d-%d)\n", i, sectionIdx, name, cps[i], cps[i+1])
				continue
			}
			
			storyText := string(fullRunes[cpStart:cpEnd])
			has01 := strings.ContainsRune(storyText, 0x01)
			has08 := strings.ContainsRune(storyText, 0x08)
			has13 := strings.ContainsRune(storyText, 0x13)
			
			fmt.Printf("  [%d] sec=%d %s (cp %d-%d, len=%d) img=%v drawn=%v field=%v\n",
				i, sectionIdx, name, cps[i], cps[i+1], storyLen, has01, has08, has13)
			
			// Show content
			runes := []rune(storyText)
			display := ""
			for _, r := range runes {
				switch {
				case r == 0x01:
					display += "[IMG]"
				case r == 0x08:
					display += "[DRAWN]"
				case r == 0x13:
					display += "[FBEGIN]"
				case r == 0x14:
					display += "[FSEP]"
				case r == 0x15:
					display += "[FEND]"
				case r == '\t':
					display += "[TAB]"
				case r == '\r':
					display += "[CR]"
				case r < 0x20:
					display += fmt.Sprintf("[%02X]", r)
				default:
					display += string(r)
				}
			}
			if len(display) > 300 {
				display = display[:300] + "..."
			}
			fmt.Printf("    %s\n", display)
		}
	}
	
	// Parse PlcSpaHdr (shapes in headers/footers)
	if lcbPlcSpaHdr > 0 {
		spaData := tableData[fcPlcSpaHdr : fcPlcSpaHdr+lcbPlcSpaHdr]
		n := (lcbPlcSpaHdr - 4) / 30
		fmt.Printf("\nPlcSpaHdr: %d shapes\n", n)
		for i := uint32(0); i < n; i++ {
			cp := binary.LittleEndian.Uint32(spaData[i*4:])
			spaOff := (n+1)*4 + i*26
			if spaOff+26 > uint32(len(spaData)) {
				break
			}
			spid := binary.LittleEndian.Uint32(spaData[spaOff:])
			fmt.Printf("  Shape[%d]: cp=%d spid=%d\n", i, cp, spid)
		}
	}
}

func readFatChain(data []byte, fat map[uint32]uint32, startSector, size, sectorSize uint32) []byte {
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

func extractPieceText(wdData, tableData []byte, fcClx, lcbClx uint32) string {
	if lcbClx == 0 {
		return ""
	}
	clxData := tableData[fcClx : fcClx+lcbClx]
	pos := 0
	for pos < len(clxData) {
		if clxData[pos] == 0x02 {
			pos++
			break
		} else if clxData[pos] == 0x01 {
			pos++
			if pos+2 > len(clxData) {
				return ""
			}
			cb := int(binary.LittleEndian.Uint16(clxData[pos:]))
			pos += 2 + cb
		} else {
			pos++
		}
	}
	if pos+4 > len(clxData) {
		return ""
	}
	lcbPlcPcd := binary.LittleEndian.Uint32(clxData[pos:])
	pos += 4
	plcPcd := clxData[pos : pos+int(lcbPlcPcd)]
	n := (int(lcbPlcPcd) - 4) / 12
	
	var result []rune
	for i := 0; i < n; i++ {
		cpStart := binary.LittleEndian.Uint32(plcPcd[i*4:])
		cpEnd := binary.LittleEndian.Uint32(plcPcd[(i+1)*4:])
		pcdOff := (n+1)*4 + i*8
		fc := binary.LittleEndian.Uint32(plcPcd[pcdOff+2:])
		isUnicode := (fc & 0x40000000) == 0
		realFC := fc & 0x3FFFFFFF
		charCount := cpEnd - cpStart
		if isUnicode {
			for j := uint32(0); j < charCount; j++ {
				off := realFC + j*2
				if int(off+2) > len(wdData) {
					break
				}
				ch := binary.LittleEndian.Uint16(wdData[off:])
				result = append(result, rune(ch))
			}
		} else {
			for j := uint32(0); j < charCount; j++ {
				off := realFC/2 + j
				if int(off) >= len(wdData) {
					break
				}
				result = append(result, rune(wdData[off]))
			}
		}
	}
	return string(result)
}
