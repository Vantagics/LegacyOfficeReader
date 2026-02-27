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

	if dataDir == nil {
		fmt.Println("No Data stream")
		return
	}

	dReader, _ := adaptor.OpenObject(dataDir, root)
	dSize := binary.LittleEndian.Uint32(dataDir.StreamSize[:])
	data := make([]byte, dSize)
	dReader.Read(data)

	fmt.Printf("Data stream size: %d bytes\n\n", len(data))

	// Scan for SpContainers
	fmt.Println("=== SpContainers in Data stream ===")
	for i := 0; i+8 <= len(data); i++ {
		verInst := binary.LittleEndian.Uint16(data[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(data[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(data[i+4 : i+8])

		if recType == 0xF004 && recVer == 0xF && recLen > 0 && uint32(i)+8+recLen <= uint32(len(data)) {
			// Parse SpContainer for pib
			spid, pib := parseSpContainerForPib(data, uint32(i)+8, uint32(i)+8+recLen)

			// Check PICF header before SpContainer
			picfInfo := ""
			for _, hdrSize := range []int{68, 44} {
				pOff := i - hdrSize
				if pOff >= 0 && pOff+6 <= len(data) {
					cb := int(binary.LittleEndian.Uint16(data[pOff+4:]))
					if cb == hdrSize && pOff+cb == i {
						picfInfo = fmt.Sprintf(" PICF@%d(cbHeader=%d)", pOff, cb)
					}
				}
			}

			fmt.Printf("SpContainer@%d: len=%d, spid=%d, pib=%d%s\n", i, recLen, spid, pib, picfInfo)

			i += int(7 + recLen)
		}
	}

	// Check the specific PicLocation values from the inline images
	fmt.Println("\n=== Checking PicLocation offsets ===")
	for _, offset := range []int{4085, 111211, 457622} {
		if offset >= 0 && offset+68 <= len(data) {
			cbHeader := binary.LittleEndian.Uint16(data[offset+4:])
			fmt.Printf("PicLoc %d: cbHeader=%d", offset, cbHeader)
			// Check what's at offset + cbHeader
			spOff := offset + int(cbHeader)
			if spOff+8 <= len(data) {
				vi := binary.LittleEndian.Uint16(data[spOff:])
				rt := binary.LittleEndian.Uint16(data[spOff+2:])
				rl := binary.LittleEndian.Uint32(data[spOff+4:])
				fmt.Printf(", at offset+cbHeader: verInst=0x%04X, recType=0x%04X, recLen=%d", vi, rt, rl)
				if rt == 0xF004 && (vi&0x0F) == 0xF {
					_, pib := parseSpContainerForPib(data, uint32(spOff)+8, uint32(spOff)+8+rl)
					fmt.Printf(" → SpContainer, pib=%d (BSE=%d)", pib, pib-1)
				}
			}
			fmt.Println()
		}
	}
}

func parseSpContainerForPib(data []byte, start, end uint32) (spid, pib uint32) {
	offset := start
	for offset+8 <= end {
		verInst := binary.LittleEndian.Uint16(data[offset:])
		recType := binary.LittleEndian.Uint16(data[offset+2:])
		recLen := binary.LittleEndian.Uint32(data[offset+4:])
		inst := verInst >> 4

		childEnd := offset + 8 + recLen
		if childEnd > end {
			childEnd = end
		}
		recData := data[offset+8 : childEnd]

		if recType == 0xF00A && len(recData) >= 8 {
			spid = binary.LittleEndian.Uint32(recData[0:])
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
				if pid == 0x0104 { // pib
					pib = propVal
				}
			}
		}
		offset = childEnd
	}
	return
}
