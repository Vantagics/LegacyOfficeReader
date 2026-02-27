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

	// Parse SpContainer records from Data stream
	fmt.Println("=== SpContainer records in Data stream ===")
	for i := uint32(0); i+8 <= uint32(len(dData)); i++ {
		verInst := binary.LittleEndian.Uint16(dData[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(dData[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(dData[i+4 : i+8])

		if recType == 0xF004 && recVer == 0xF && recLen > 0 && i+8+recLen <= uint32(len(dData)) {
			fmt.Printf("\nSpContainer at offset %d, len=%d\n", i, recLen)
			parseSpContainer(dData, i+8, i+8+recLen)
			i += 7 + recLen
		}
	}
}

func parseSpContainer(data []byte, start, end uint32) {
	for offset := start; offset+8 <= end; {
		verInst := binary.LittleEndian.Uint16(data[offset : offset+2])
		recVer := verInst & 0x0F
		recInst := verInst >> 4
		recType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
		recLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])

		childEnd := offset + 8 + recLen
		if childEnd > end {
			break
		}

		switch recType {
		case 0xF00A: // Sp
			if recLen >= 8 {
				spid := binary.LittleEndian.Uint32(data[offset+8:])
				flags := binary.LittleEndian.Uint32(data[offset+12:])
				fmt.Printf("  Sp: spid=%d flags=0x%08X inst=0x%03X\n", spid, flags, recInst)
			}
		case 0xF00B: // Opt
			fmt.Printf("  Opt: len=%d inst=0x%03X\n", recLen, recInst)
			parseOpt(data, offset+8, childEnd)
		case 0xF010: // ClientAnchor
			if recLen >= 4 {
				anchor := binary.LittleEndian.Uint32(data[offset+8:])
				fmt.Printf("  ClientAnchor: value=%d\n", anchor)
			}
		case 0xF011: // ClientData
			fmt.Printf("  ClientData: len=%d\n", recLen)
		case 0xF007: // BSE (embedded)
			fmt.Printf("  BSE: len=%d\n", recLen)
		case 0xF122: // TertiaryOpt
			fmt.Printf("  TertiaryOpt: len=%d\n", recLen)
		default:
			if recType >= 0xF01A && recType <= 0xF02A {
				fmt.Printf("  Blip: type=0x%04X len=%d\n", recType, recLen)
			} else {
				fmt.Printf("  Record: type=0x%04X len=%d ver=0x%X\n", recType, recLen, recVer)
			}
		}

		offset = childEnd
	}
}

func parseOpt(data []byte, start, end uint32) {
	// Opt record: array of FOPTE entries (6 bytes each: pid(2) + value(4))
	// followed by complex data
	for offset := start; offset+6 <= end; {
		pidRaw := binary.LittleEndian.Uint16(data[offset:])
		pid := pidRaw & 0x3FFF
		fBid := (pidRaw >> 14) & 1
		fComplex := (pidRaw >> 15) & 1
		value := binary.LittleEndian.Uint32(data[offset+2:])

		// Key properties:
		// 0x0104 (260) = pib (picture index into BSE array, 1-based)
		// 0x0181 (385) = fillType
		// 0x01BF (447) = fNoFillHitTest
		switch pid {
		case 260: // pib
			fmt.Printf("    pib=%d (BSE index, 1-based) fBid=%d fComplex=%d\n", value, fBid, fComplex)
		case 261: // pibName
			fmt.Printf("    pibName value=%d\n", value)
		case 262: // pibFlags
			fmt.Printf("    pibFlags=0x%08X\n", value)
		case 385: // fillType
			fmt.Printf("    fillType=%d\n", value)
		case 447: // fNoFillHitTest
			fmt.Printf("    fNoFillHitTest=0x%08X\n", value)
		default:
			if fBid != 0 || pid == 260 {
				fmt.Printf("    pid=%d value=%d fBid=%d fComplex=%d\n", pid, value, fBid, fComplex)
			}
		}

		offset += 6
	}
}
