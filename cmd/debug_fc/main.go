package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/helpers"
)

func main() {
	adaptor, err := cfb.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer adaptor.CloseFile()

	var root, wordDoc, table1 *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "WordDocument":
			wordDoc = dir
		case "1Table":
			table1 = dir
		}
	}

	wReader, _ := adaptor.OpenObject(wordDoc, root)
	wSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wData := make([]byte, wSize)
	wReader.Read(wData)

	tReader, _ := adaptor.OpenObject(table1, root)
	tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tData := make([]byte, tSize)
	tReader.Read(tData)

	// Parse FIB
	offset := 0x20
	csw := binary.LittleEndian.Uint16(wData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wData[offset:])
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wData[offset:])
	offset += 2

	readFcLcb := func(index int) uint32 {
		if int(cbRgFcLcb) <= index {
			return 0
		}
		off := offset + index*4
		return binary.LittleEndian.Uint32(wData[off:])
	}

	fcClx := readFcLcb(66)
	lcbClx := readFcLcb(67)

	// Parse piece table
	clxData := tData[fcClx : fcClx+lcbClx]
	pos := uint32(0)
	for pos < uint32(len(clxData)) {
		if clxData[pos] == 0x01 {
			prcSize := binary.LittleEndian.Uint16(clxData[pos+1:])
			pos += 3 + uint32(prcSize)
			continue
		}
		if clxData[pos] == 0x02 {
			pos++
			break
		}
		break
	}
	plcPcdLen := binary.LittleEndian.Uint32(clxData[pos:])
	pos += 4
	plcPcdData := clxData[pos : pos+plcPcdLen]
	n := (plcPcdLen - 4) / 12

	cps := make([]uint32, n+1)
	for i := uint32(0); i <= n; i++ {
		cps[i] = binary.LittleEndian.Uint32(plcPcdData[i*4:])
	}

	// For CPs 6, 9, 12, 15, 108, find the FC and read the raw bytes
	checkCPs := []uint32{6, 9, 12, 15, 108}
	for _, cp := range checkCPs {
		// Find which piece contains this CP
		for i := uint32(0); i < n; i++ {
			if cp >= cps[i] && cp < cps[i+1] {
				pdStart := (n+1)*4 + i*8
				fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])
				isUnicode := fc&0x40000000 == 0
				var actualOffset uint32
				if isUnicode {
					actualOffset = fc
				} else {
					actualOffset = (fc & ^uint32(0x40000000)) >> 1
				}

				cpOffset := cp - cps[i]
				var byteOffset uint32
				if isUnicode {
					byteOffset = actualOffset + cpOffset*2
				} else {
					byteOffset = actualOffset + cpOffset
				}

				if isUnicode {
					// Read 2 bytes (UTF-16LE)
					if int(byteOffset)+2 <= len(wData) {
						lo := wData[byteOffset]
						hi := wData[byteOffset+1]
						char := helpers.DecodeUTF16LE([]byte{lo, hi})
						r := []rune(char)
						if len(r) > 0 && r[0] < 0x20 {
							fmt.Printf("CP %d: FC=%d bytes=[%02X %02X] = \\x%02X (unicode)\n", cp, byteOffset, lo, hi, r[0])
						} else {
							fmt.Printf("CP %d: FC=%d bytes=[%02X %02X] = %q (unicode)\n", cp, byteOffset, lo, hi, char)
						}
					}
				} else {
					if int(byteOffset) < len(wData) {
						b := wData[byteOffset]
						fmt.Printf("CP %d: FC=%d byte=[%02X] (ANSI)\n", cp, byteOffset, b)
					}
				}
				break
			}
		}
	}
}
