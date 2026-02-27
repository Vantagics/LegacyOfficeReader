package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"example.com/officeconv/cfb"
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

func (rh recordHeader) recVer() uint16  { return rh.recVerAndInstance & 0x0F }
func (rh recordHeader) recInstance() uint16 { return rh.recVerAndInstance >> 4 }

const (
	rtSlideListWithText = 0x0FF0
	rtSlidePersistAtom  = 0x03F3
	rtUserEdit          = 0x0FF5
	rtPersistDirAtom    = 0x1772
	rtCurrentUserAtom   = 0x0FF6