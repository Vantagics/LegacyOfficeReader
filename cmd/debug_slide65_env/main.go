package main

import (
	"encoding/binary"
	"fmt"

	"github.com/shakinm/xlsReader/cfb"
)

const recordHeaderSize = 8

type recordHeader struct {
	recVerAndInstance uint16
	recType           uint16
	recLen            uint32
}

func readRecordHeader(data []byte, offset uint32) (recordHeader, error) {
	if uint32(len(data)) < offset+recordHeaderSize {
		return recordHeader{}, fmt.Errorf("not enough data")
	}
	return recordHeader{
		recVerAndInstance: binary.LittleEndian.Uint16(data[offset:]),
		recType:           binary.LittleEndian.Uint16(data[offset+2:]),
		recLen:            binary.LittleEndian.Uint32(data[offset+4:]),
	}, nil
}

func (rh recordHeader) recVer() uint16     { return rh.recVerAndInstance & 0x0F }
func (rh recordHeader) recInstance() uint16 { return rh.recVerAndInstance >> 4 }

func main() {
	adaptor, err := cfb.OpenFile("testfie/test.ppt")
	if err != nil {
		panic(err)
	}
	defer adaptor.CloseFile()

	var root *cfb.Directory
	var pptDoc *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "PowerPoint Document":
			pptDoc = dir
		}
	}

	reader, _ := adaptor.OpenObject(pptDoc, root)
	size := binary.LittleEndian.Uint32(pptDoc.StreamSize[:])
	data := make([]byte, size)
	reader.Read(data)

	dataLen := uint32(len(data))
	offset := uint32(0)

	for offset+recordHeaderSize <= dataLen {
		rh, err := readRecordHeader(data, offset)
		if err != nil {
			break
		}
		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > dataLen {
			break
		}

		// Environment container = 0x03F2 (1010)
		if rh.recType == 0x03F2 {
			fmt.Printf("Found Environment at offset %d, len=%d\n", offset, rh.recLen)
			// Scan inside
			pos := recDataStart
			for pos+recordHeaderSize <= recDataEnd {
				sub, err := readRecordHeader(data, pos)
				if err != nil {
					break
				}
				subDataStart := pos + recordHeaderSize
				subDataEnd := subDataStart + sub.recLen
				if subDataEnd > recDataEnd {
					break
				}

				fmt.Printf("  Sub: type=0x%04X ver=%d inst=%d len=%d\n",
					sub.recType, sub.recVer(), sub.recInstance(), sub.recLen)

				// TextMasterStyleAtom = 0x0FA3
				if sub.recType == 0x0FA3 {
					fmt.Printf("    TextMasterStyleAtom instance=%d (textType=%d)\n",
						sub.recInstance(), sub.recInstance())
					// Parse first 2 bytes = numLevels
					if sub.recLen >= 2 {
						numLevels := binary.LittleEndian.Uint16(data[subDataStart : subDataStart+2])
						fmt.Printf("    numLevels=%d\n", numLevels)
					}
				}

				if sub.recVer() == 0xF {
					pos = subDataStart
				} else {
					pos = subDataEnd
				}
			}
		}

		if rh.recVer() == 0xF {
			offset = recDataStart
		} else {
			offset = recDataEnd
		}
	}
}
