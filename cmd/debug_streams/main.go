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

	var root *cfb.Directory
	fmt.Println("=== CFB Directory Entries ===")
	for _, dir := range adaptor.GetDirs() {
		name := dir.Name()
		size := binary.LittleEndian.Uint32(dir.StreamSize[:])
		fmt.Printf("  %q size=%d\n", name, size)
		if name == "Root Entry" {
			root = dir
		}
	}

	// Try to read each stream and show first bytes
	for _, dir := range adaptor.GetDirs() {
		name := dir.Name()
		size := binary.LittleEndian.Uint32(dir.StreamSize[:])
		if size == 0 || name == "Root Entry" {
			continue
		}

		reader, err := adaptor.OpenObject(dir, root)
		if err != nil {
			fmt.Printf("  Cannot open %q: %v\n", name, err)
			continue
		}

		data := make([]byte, size)
		n, _ := reader.Read(data)
		data = data[:n]

		fmt.Printf("\n=== Stream %q (%d bytes) ===\n", name, n)

		// For Data stream, scan for OfficeArt records
		if name == "Data" || name == "Pictures" {
			scanRecords(data, name)
		}
	}
}

func scanRecords(data []byte, name string) {
	fmt.Printf("Scanning %s for OfficeArt records...\n", name)
	for i := 0; i+8 <= len(data); i++ {
		recType := binary.LittleEndian.Uint16(data[i+2 : i+4])
		verInst := binary.LittleEndian.Uint16(data[i : i+2])
		recVer := verInst & 0x0F
		recLen := binary.LittleEndian.Uint32(data[i+4 : i+8])

		// Only look for known OfficeArt types
		if recType < 0xF000 || recType > 0xF12A {
			continue
		}
		if recLen == 0 || uint32(i)+8+recLen > uint32(len(data)) {
			continue
		}

		typeName := ""
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
			typeName = "Dgg"
		case 0xF007:
			typeName = "BSE"
		case 0xF008:
			typeName = "Dg"
		case 0xF009:
			typeName = "Spgr"
		case 0xF00A:
			typeName = "Sp"
		case 0xF00B:
			typeName = "Opt"
		case 0xF010:
			typeName = "ClientAnchor"
		case 0xF011:
			typeName = "ClientData"
		case 0xF01A:
			typeName = "BlipEMF"
		case 0xF01B:
			typeName = "BlipWMF"
		case 0xF01D:
			typeName = "BlipJPEG"
		case 0xF01E:
			typeName = "BlipPNG"
		case 0xF01F:
			typeName = "BlipDIB"
		case 0xF11E:
			typeName = "SplitMenuColors"
		case 0xF122:
			typeName = "TertiaryOpt"
		default:
			typeName = fmt.Sprintf("Unknown(0x%04X)", recType)
		}

		fmt.Printf("  offset=%d ver=0x%X type=0x%04X(%s) len=%d\n",
			i, recVer, recType, typeName, recLen)

		// If it's a BSE, show the blip type inside
		if recType == 0xF007 && recVer == 0x2 {
			if i+8+36+8 <= len(data) {
				blipType := data[i+8] // btWin32
				fmt.Printf("    BSE btWin32=%d\n", blipType)
				// Check for embedded blip
				blipOff := uint32(i) + 8 + 36
				if blipOff+8 <= uint32(len(data)) {
					blipRecType := binary.LittleEndian.Uint16(data[blipOff+2 : blipOff+4])
					blipRecLen := binary.LittleEndian.Uint32(data[blipOff+4 : blipOff+8])
					fmt.Printf("    Embedded blip: type=0x%04X len=%d\n", blipRecType, blipRecLen)
				}
			}
		}
	}
}
