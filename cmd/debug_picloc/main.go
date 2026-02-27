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

	var root, wordDoc, table1, dataDir *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "WordDocument":
			wordDoc = dir
		case "1Table":
			table1 = dir
		case "Data":
			dataDir = dir
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

	var dData []byte
	if dataDir != nil {
		dReader, _ := adaptor.OpenObject(dataDir, root)
		dSize := binary.LittleEndian.Uint32(dataDir.StreamSize[:])
		dData = make([]byte, dSize)
		dReader.Read(dData)
	}

	// Parse FIB to get CHPX info
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

	fcPlcfBteChpx := readFcLcb(24)
	lcbPlcfBteChpx := readFcLcb(25)

	// Parse CHPX to find sprmCPicLocation values
	fmt.Println("=== Scanning CHPX for sprmCPicLocation (0x6A03) ===")
	plcData := tData[fcPlcfBteChpx : fcPlcfBteChpx+lcbPlcfBteChpx]
	n := (lcbPlcfBteChpx - 4) / 8

	for i := uint32(0); i < n; i++ {
		pnOffset := (n+1)*4 + i*4
		pn := binary.LittleEndian.Uint32(plcData[pnOffset:])
		pageOffset := pn * 512
		if pageOffset+512 > uint32(len(wData)) {
			continue
		}
		page := wData[pageOffset : pageOffset+512]
		crun := int(page[511])
		fcArraySize := (crun + 1) * 4
		rgbStart := fcArraySize

		for j := 0; j < crun; j++ {
			if rgbStart+j >= 511 {
				break
			}
			rgb := int(page[rgbStart+j])
			if rgb == 0 {
				continue
			}
			chpxPos := rgb * 2
			if chpxPos >= 512 {
				continue
			}
			cb := int(page[chpxPos])
			if chpxPos+1+cb > 512 {
				continue
			}
			sprmData := page[chpxPos+1 : chpxPos+1+cb]

			// Scan for sprmCPicLocation (0x6A03)
			pos := 0
			for pos+2 <= len(sprmData) {
				opcode := binary.LittleEndian.Uint16(sprmData[pos:])
				pos += 2
				opSize := sprmOperandSize(opcode)
				if opSize == -1 {
					if pos >= len(sprmData) {
						break
					}
					opSize = int(sprmData[pos])
					pos++
				}
				if pos+opSize > len(sprmData) {
					break
				}
				if opcode == 0x6A03 && opSize >= 4 {
					picLoc := int32(binary.LittleEndian.Uint32(sprmData[pos:]))
					fcStart := binary.LittleEndian.Uint32(page[j*4:])
					fcEnd := binary.LittleEndian.Uint32(page[(j+1)*4:])
					fmt.Printf("  FC[%d-%d] sprmCPicLocation=%d\n", fcStart, fcEnd, picLoc)

					// Check what's at that offset in the Data stream
					if picLoc >= 0 && int(picLoc)+68 <= len(dData) {
						// PICF header is 68 bytes, then OfficeArt data
						lcb := binary.LittleEndian.Uint32(dData[picLoc:])
						cbHeader := binary.LittleEndian.Uint16(dData[picLoc+4:])
						fmt.Printf("    PICF: lcb=%d cbHeader=%d\n", lcb, cbHeader)

						// After PICF header, look for SpContainer
						picfEnd := int(picLoc) + int(cbHeader)
						if picfEnd+8 <= len(dData) {
							verInst := binary.LittleEndian.Uint16(dData[picfEnd:])
							recVer := verInst & 0x0F
							recType := binary.LittleEndian.Uint16(dData[picfEnd+2:])
							recLen := binary.LittleEndian.Uint32(dData[picfEnd+4:])
							fmt.Printf("    After PICF: ver=0x%X type=0x%04X len=%d\n", recVer, recType, recLen)

							if recType == 0xF004 && recVer == 0xF {
								// Parse SpContainer for pib
								spid, pib := parseSpContainerForPib2(dData, uint32(picfEnd)+8, uint32(picfEnd)+8+recLen)
								fmt.Printf("    SpContainer: SPID=%d pib=%d\n", spid, pib)
							}
						}
					}
				}
				pos += opSize
			}
		}
	}
}

func sprmOperandSize(opcode uint16) int {
	spra := (opcode >> 13) & 0x07
	switch spra {
	case 0:
		return 1
	case 1:
		return 1
	case 2:
		return 2
	case 3:
		return 4
	case 4:
		return 2
	case 5:
		return 2
	case 6:
		return -1
	case 7:
		return 3
	}
	return 0
}

func parseSpContainerForPib2(data []byte, offset, limit uint32) (spid, pib uint32) {
	for offset+8 <= limit {
		verInst := binary.LittleEndian.Uint16(data[offset : offset+2])
		recVer := verInst & 0x0F
		recInst := verInst >> 4
		recType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
		recLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])
		childEnd := offset + 8 + recLen
		if childEnd > limit {
			break
		}
		if recType == 0xF00A && recLen >= 8 {
			spid = binary.LittleEndian.Uint32(data[offset+8:])
		}
		if recType == 0xF00B && recLen > 0 {
			nProps := recInst
			propOff := offset + 8
			for p := uint16(0); p < nProps && propOff+6 <= childEnd; p++ {
				propID := binary.LittleEndian.Uint16(data[propOff:])
				propVal := binary.LittleEndian.Uint32(data[propOff+2:])
				pid := propID & 0x3FFF
				if pid == 260 {
					pib = propVal
				}
				propOff += 6
			}
		}
		if recVer == 0xF {
			s, p := parseSpContainerForPib2(data, offset+8, childEnd)
			if s != 0 {
				spid = s
			}
			if p != 0 {
				pib = p
			}
		}
		offset = childEnd
	}
	return
}
