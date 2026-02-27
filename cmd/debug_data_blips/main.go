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

	var root, dataDir *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "Data":
			dataDir = dir
		}
	}

	dReader, _ := adaptor.OpenObject(dataDir, root)
	dSize := binary.LittleEndian.Uint32(dataDir.StreamSize[:])
	dData := make([]byte, dSize)
	dReader.Read(dData)

	fmt.Printf("Data stream: %d bytes\n", len(dData))

	// Check blip records at the foDelay offsets from BSE entries
	foDelays := []struct {
		idx     int
		offset  uint32
		size    uint32
		btWin32 byte
	}{
		{0, 42034, 109874, 6},   // DIB
		{1, 151908, 346364, 2},  // EMF
		{2, 498272, 5088, 6},    // DIB
		{3, 503360, 10799, 6},   // DIB
		{4, 514159, 209536, 2},  // EMF
		{5, 723695, 359312, 2},  // EMF
		{6, 1083007, 34982, 6},  // DIB
		{7, 1117989, 42582, 6},  // DIB
	}

	for _, fd := range foDelays {
		fmt.Printf("\nBSE[%d] foDelay=%d size=%d btWin32=%d:\n", fd.idx, fd.offset, fd.size, fd.btWin32)
		if fd.offset+8 > uint32(len(dData)) {
			fmt.Printf("  OFFSET OUT OF BOUNDS (data stream is %d bytes)\n", len(dData))
			continue
		}

		// Check if there's an OfficeArt record header at this offset
		verInst := binary.LittleEndian.Uint16(dData[fd.offset : fd.offset+2])
		recVer := verInst & 0x0F
		recInst := verInst >> 4
		recType := binary.LittleEndian.Uint16(dData[fd.offset+2 : fd.offset+4])
		recLen := binary.LittleEndian.Uint32(dData[fd.offset+4 : fd.offset+8])

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
		case 0xF004:
			typeName = "SpContainer"
		case 0xF007:
			typeName = "BSE"
		}

		fmt.Printf("  Record: ver=0x%X inst=0x%03X type=0x%04X(%s) len=%d\n",
			recVer, recInst, recType, typeName, recLen)

		// Also check a few bytes before the offset for container records
		if fd.offset >= 8 {
			prevType := binary.LittleEndian.Uint16(dData[fd.offset-6 : fd.offset-4])
			prevLen := binary.LittleEndian.Uint32(dData[fd.offset-4 : fd.offset])
			fmt.Printf("  Prev record at offset %d: type=0x%04X len=%d\n",
				fd.offset-8, prevType, prevLen)
		}
	}
}
