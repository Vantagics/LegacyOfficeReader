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

	var wordDoc, table0, table1, root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "WordDocument":
			wordDoc = dir
		case "0Table":
			table0 = dir
		case "1Table":
			table1 = dir
		case "Root Entry":
			root = dir
		}
	}

	wordDocReader, _ := adaptor.OpenObject(wordDoc, root)
	wordDocSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wordDocData := make([]byte, wordDocSize)
	wordDocReader.Read(wordDocData)

	flags := binary.LittleEndian.Uint16(wordDocData[0x0A:])
	fWhichTblStm := (flags >> 9) & 1
	var tableDir *cfb.Directory
	if fWhichTblStm == 1 {
		tableDir = table1
	} else {
		tableDir = table0
	}

	tableReader, _ := adaptor.OpenObject(tableDir, root)
	tableSize := binary.LittleEndian.Uint32(tableDir.StreamSize[:])
	tableData := make([]byte, tableSize)
	tableReader.Read(tableData)

	// Navigate FIB to find Clx
	offset := 0x20
	csw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2

	readFcLcb := func(index int) uint32 {
		if int(cbRgFcLcb) <= index {
			return 0
		}
		off := offset + index*4
		if off+4 > len(wordDocData) {
			return 0
		}
		return binary.LittleEndian.Uint32(wordDocData[off:])
	}

	fcClx := readFcLcb(66)
	lcbClx := readFcLcb(67)
	fmt.Printf("fcClx=%d lcbClx=%d\n", fcClx, lcbClx)

	// Parse Clx
	clxData := tableData[fcClx : fcClx+lcbClx]
	pos := uint32(0)
	for pos < uint32(len(clxData)) {
		typeByte := clxData[pos]
		if typeByte == 0x01 {
			prcSize := binary.LittleEndian.Uint16(clxData[pos+1:])
			pos += 3 + uint32(prcSize)
			continue
		}
		if typeByte == 0x02 {
			pos++
			break
		}
		fmt.Printf("Unknown type byte: 0x%02X at offset %d\n", typeByte, pos)
		break
	}

	plcPcdLen := binary.LittleEndian.Uint32(clxData[pos:])
	pos += 4
	plcPcdData := clxData[pos : pos+plcPcdLen]

	n := (plcPcdLen - 4) / 12
	fmt.Printf("Number of pieces: %d\n", n)

	cps := make([]uint32, n+1)
	for i := uint32(0); i <= n; i++ {
		cps[i] = binary.LittleEndian.Uint32(plcPcdData[i*4:])
	}

	fmt.Printf("CP range: [%d, %d]\n", cps[0], cps[n])
	for i := uint32(0); i < n; i++ {
		cpArraySize := (n + 1) * 4
		pdStart := cpArraySize + i*8
		fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])
		isUnicode := fc&0x40000000 == 0
		fmt.Printf("  Piece[%d]: CP[%d-%d] fc=0x%08X unicode=%v\n", i, cps[i], cps[i+1], fc, isUnicode)
	}
}
