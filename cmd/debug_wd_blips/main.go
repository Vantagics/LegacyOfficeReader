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

	var root, wordDoc *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "WordDocument":
			wordDoc = dir
		}
	}

	wReader, _ := adaptor.OpenObject(wordDoc, root)
	wSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wData := make([]byte, wSize)
	wReader.Read(wData)

	fmt.Printf("WordDocument stream: %d bytes\n", len(wData))

	foDelays := []struct {
		idx     int
		offset  uint32
		size    uint32
		btWin32 byte
	}{
		{0, 42034, 109874, 6},
		{1, 151908, 346364, 2},
		{2, 498272, 5088, 6},
		{3, 503360, 10799, 6},
		{4, 514159, 209536, 2},
		{5, 723695, 359312, 2},
		{6, 1083007, 34982, 6},
		{7, 1117989, 42582, 6},
	}

	for _, fd := range foDelays {
		fmt.Printf("\nBSE[%d] foDelay=%d size=%d btWin32=%d:\n", fd.idx, fd.offset, fd.size, fd.btWin32)
		if fd.offset+8 > uint32(len(wData)) {
			fmt.Printf("  OFFSET OUT OF BOUNDS\n")
			continue
		}

		verInst := binary.LittleEndian.Uint16(wData[fd.offset : fd.offset+2])
		recVer := verInst & 0x0F
		recInst := verInst >> 4
		recType := binary.LittleEndian.Uint16(wData[fd.offset+2 : fd.offset+4])
		recLen := binary.LittleEndian.Uint32(wData[fd.offset+4 : fd.offset+8])

		typeName := "unknown"
		switch recType {
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
		}

		fmt.Printf("  Record: ver=0x%X inst=0x%03X type=0x%04X(%s) len=%d\n",
			recVer, recInst, recType, typeName, recLen)

		if recType >= 0xF01A && recType <= 0xF02A {
			fmt.Printf("  FOUND BLIP! Total consumed: %d bytes\n", 8+recLen)
			// Verify the data is reasonable
			if fd.offset+8+recLen <= uint32(len(wData)) {
				fmt.Printf("  Data available: YES\n")
			} else {
				fmt.Printf("  Data available: NO (truncated)\n")
			}
		}
	}
}
