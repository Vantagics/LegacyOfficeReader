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
	var table1 *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "1Table":
			table1 = dir
		}
	}

	tReader, _ := adaptor.OpenObject(table1, root)
	tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tData := make([]byte, tSize)
	tReader.Read(tData)

	// BSE structure per [MS-ODRAW]:
	// Byte 0: btWin32 (1 byte) - blip type for Windows
	// Byte 1: btMacOS (1 byte) - blip type for Mac
	// Bytes 2-17: rgbUid (16 bytes) - UID of the blip
	// Byte 18-19: tag (uint16)
	// Bytes 20-23: size (uint32) - size of the blip in the stream
	// Bytes 24-27: cRef (uint32) - reference count
	// Bytes 28-31: foDelay (uint32) - offset in the delay stream (Data stream)
	// Byte 32: usage (1 byte)
	// Byte 33: cbName (1 byte) - length of name
	// Bytes 34-35: unused (2 bytes)

	// The BSE entries start at offset 9468 in the Table stream
	bseOffsets := []uint32{9468, 9512, 9556, 9600, 9644, 9688, 9732, 9776}

	for i, off := range bseOffsets {
		if off+8+36 > uint32(len(tData)) {
			break
		}
		bseData := tData[off+8 : off+8+36] // Skip 8-byte record header

		btWin32 := bseData[0]
		btMacOS := bseData[1]
		// rgbUid at bytes 2-17
		tag := binary.LittleEndian.Uint16(bseData[18:20])
		size := binary.LittleEndian.Uint32(bseData[20:24])
		cRef := binary.LittleEndian.Uint32(bseData[24:28])
		foDelay := binary.LittleEndian.Uint32(bseData[28:32])
		usage := bseData[32]
		cbName := bseData[33]

		fmt.Printf("BSE[%d] at offset %d:\n", i, off)
		fmt.Printf("  btWin32=%d btMacOS=%d tag=%d\n", btWin32, btMacOS, tag)
		fmt.Printf("  size=%d cRef=%d foDelay=%d\n", size, cRef, foDelay)
		fmt.Printf("  usage=%d cbName=%d\n", usage, cbName)
		fmt.Printf("  rgbUid=%X\n", bseData[2:18])

		// btWin32 values: 2=EMF, 3=WMF, 4=JPEG, 5=PNG, 6=DIB, 7=TIFF
		btNames := map[byte]string{1: "ERROR", 2: "EMF", 3: "WMF", 4: "JPEG", 5: "PNG", 6: "DIB", 7: "TIFF"}
		if name, ok := btNames[btWin32]; ok {
			fmt.Printf("  Format: %s\n", name)
		}
		fmt.Println()
	}
}
