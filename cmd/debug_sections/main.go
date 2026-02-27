package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
)

func main() {
	f, _ := os.Open("testfie/test.doc")
	defer f.Close()

	adaptor, _ := cfb.OpenReader(f)
	var wordDoc, table1, root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "WordDocument":
			wordDoc = dir
		case "1Table":
			table1 = dir
		case "Root Entry":
			root = dir
		}
	}

	wdReader, _ := adaptor.OpenObject(wordDoc, root)
	wdSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wdData := make([]byte, wdSize)
	wdReader.Read(wdData)

	tReader, _ := adaptor.OpenObject(table1, root)
	tSize := binary.LittleEndian.Uint32(table1.StreamSize[:])
	tData := make([]byte, tSize)
	tReader.Read(tData)

	// Navigate FIB
	offset := 0x20
	csw := binary.LittleEndian.Uint16(wdData[offset:])
	offset += 2 + int(csw)*2
	cslw := binary.LittleEndian.Uint16(wdData[offset:])
	fibRgLwStart := offset + 2
	ccpText := binary.LittleEndian.Uint32(wdData[fibRgLwStart+3*4:])
	offset += 2 + int(cslw)*4
	cbRgFcLcb := binary.LittleEndian.Uint16(wdData[offset:])
	offset += 2

	readFcLcb := func(idx int) uint32 {
		if int(cbRgFcLcb) <= idx { return 0 }
		off := offset + idx*4
		return binary.LittleEndian.Uint32(wdData[off:])
	}

	fcPlcfSed := readFcLcb(12)
	lcbPlcfSed := readFcLcb(13)

	fmt.Printf("ccpText=%d\n", ccpText)
	fmt.Printf("fcPlcfSed=%d lcbPlcfSed=%d\n", fcPlcfSed, lcbPlcfSed)

	// Parse PlcfSed: (n+1) CPs (4 bytes each) + n SEDs (12 bytes each)
	// Total: (n+1)*4 + n*12 = 4 + 16n
	if lcbPlcfSed > 0 {
		sedData := tData[fcPlcfSed : fcPlcfSed+lcbPlcfSed]
		n := (lcbPlcfSed - 4) / 16
		fmt.Printf("PlcfSed: %d sections\n\n", n)

		for i := uint32(0); i <= n; i++ {
			cp := binary.LittleEndian.Uint32(sedData[i*4:])
			fmt.Printf("  CP[%d] = %d", i, cp)
			if i < n {
				// SED at offset (n+1)*4 + i*12
				sedOff := (n+1)*4 + i*12
				if sedOff+12 <= uint32(len(sedData)) {
					// SED: fn(2) + fcSepx(4) + fnMpr(2) + fcMpr(4)
					fcSepx := binary.LittleEndian.Uint32(sedData[sedOff+2:])
					fmt.Printf(" fcSepx=%d", fcSepx)

					// Read SEPX from WordDocument stream
					if fcSepx > 0 && fcSepx != 0xFFFFFFFF && int(fcSepx)+2 <= len(wdData) {
						cbSepx := binary.LittleEndian.Uint16(wdData[fcSepx:])
						fmt.Printf(" cbSepx=%d", cbSepx)
						if cbSepx > 0 && int(fcSepx)+2+int(cbSepx) <= len(wdData) {
							sprmData := wdData[fcSepx+2 : fcSepx+2+uint32(cbSepx)]
							// Parse sprms
							for pos := 0; pos+2 <= len(sprmData); {
								op := binary.LittleEndian.Uint16(sprmData[pos:])
								pos += 2
								opSize := 1
								switch (op >> 13) & 0x07 {
								case 0, 1: opSize = 1
								case 2: opSize = 2
								case 3: opSize = 4
								case 4, 5: opSize = 2
								case 6: opSize = 0 // variable
								case 7: opSize = 3
								}
								if op == 0x3009 { // sprmSBkc
									if pos < len(sprmData) {
										bkc := sprmData[pos]
										bkcNames := map[byte]string{0: "continuous", 1: "newColumn", 2: "newPage", 3: "evenPage", 4: "oddPage"}
										name := bkcNames[bkc]
										if name == "" { name = fmt.Sprintf("unknown(%d)", bkc) }
										fmt.Printf(" bkc=%s", name)
									}
								}
								if pos+opSize > len(sprmData) { break }
								pos += opSize
							}
						}
					}
				}
			}
			fmt.Println()
		}
	}
}
