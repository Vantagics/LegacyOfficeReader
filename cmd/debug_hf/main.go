package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
)

const recordHeaderSize = 8

type recordHeader struct {
	recVerInst uint16
	recType    uint16
	recLen     uint32
}

func (rh recordHeader) recVer() uint8     { return uint8(rh.recVerInst & 0x0F) }
func (rh recordHeader) recInstance() uint16 { return rh.recVerInst >> 4 }

func readRecordHeader(data []byte, offset uint32) (recordHeader, error) {
	if offset+8 > uint32(len(data)) {
		return recordHeader{}, fmt.Errorf("out of bounds")
	}
	return recordHeader{
		recVerInst: binary.LittleEndian.Uint16(data[offset : offset+2]),
		recType:    binary.LittleEndian.Uint16(data[offset+2 : offset+4]),
		recLen:     binary.LittleEndian.Uint32(data[offset+4 : offset+8]),
	}, nil
}

func main() {
	adaptor, err := cfb.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer adaptor.CloseFile()

	var pptDoc, root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "PowerPoint Document":
			pptDoc = dir
		case "Root Entry":
			root = dir
		}
	}

	pptDocReader, _ := adaptor.OpenObject(pptDoc, root)
	pptDocSize := binary.LittleEndian.Uint32(pptDoc.StreamSize[:])
	data := make([]byte, pptDocSize)
	pptDocReader.Read(data)

	dataLen := uint32(len(data))

	// Scan for HeadersFootersContainer (0x0FD9) and related records
	fmt.Println("=== Scanning for Headers/Footers records ===")
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

		// HeadersFootersContainer = 0x0FD9 (4057)
		if rh.recType == 0x0FD9 {
			fmt.Printf("HeadersFootersContainer at offset %d (len=%d, instance=%d)\n", offset, rh.recLen, rh.recInstance())
			// Parse sub-records
			pos := recDataStart
			for pos+recordHeaderSize <= recDataEnd {
				sub, err := readRecordHeader(data, pos)
				if err != nil {
					break
				}
				subStart := pos + recordHeaderSize
				subEnd := subStart + sub.recLen
				if subEnd > recDataEnd {
					break
				}

				switch sub.recType {
				case 0x0FDA: // HeadersFootersAtom
					if sub.recLen >= 4 {
						formatId := binary.LittleEndian.Uint16(data[subStart : subStart+2])
						flags := binary.LittleEndian.Uint16(data[subStart+2 : subStart+4])
						fmt.Printf("  HeadersFootersAtom: formatId=%d, flags=0x%04X\n", formatId, flags)
						fmt.Printf("    fHasDate=%v, fHasTodayDate=%v, fHasUserDate=%v\n",
							flags&0x01 != 0, flags&0x02 != 0, flags&0x04 != 0)
						fmt.Printf("    fHasSlideNumber=%v, fHasHeader=%v, fHasFooter=%v\n",
							flags&0x08 != 0, flags&0x10 != 0, flags&0x20 != 0)
					}
				case 0x0FBA: // CString (footer text, header text, date text)
					text := ""
					if sub.recLen > 0 {
						// UTF-16LE
						for i := subStart; i+1 < subEnd; i += 2 {
							ch := binary.LittleEndian.Uint16(data[i : i+2])
							if ch == 0 {
								break
							}
							text += string(rune(ch))
						}
					}
					fmt.Printf("  CString (instance=%d): %q\n", sub.recInstance(), text)
				}

				pos = subEnd
			}
		}

		// Also look for PerSlideHeadersFootersContainer (0x03F2 = 1010)
		if rh.recType == 0x03F2 {
			fmt.Printf("PerSlideHeadersFootersContainer at offset %d\n", offset)
		}

		if rh.recVer() == 0xF {
			offset = recDataStart
		} else {
			offset = recDataEnd
		}
	}
}
