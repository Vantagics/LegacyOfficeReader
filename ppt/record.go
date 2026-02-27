package ppt

import (
	"encoding/binary"
	"errors"
)

// recordHeader represents a PPT record header (8 bytes).
type recordHeader struct {
	recVerAndInstance uint16 // low 4 bits = recVer, high 12 bits = recInstance
	recType           uint16
	recLen             uint32
}

// PPT RecordType constants
const (
	rtDocument             = 0x03E8 // 1000
	rtSlideListWithText    = 0x0FF0 // 4080
	rtTextCharsAtom        = 0x0FA0 // 4000
	rtTextBytesAtom        = 0x0FA8 // 4008
	rtUserEdit             = 0x0FF5 // 4085
	rtPersistDirectoryAtom = 0x1772 // 6002
	rtCurrentUserAtom      = 0x0FF6 // 4086
	rtEndDocument          = 0x03EA // 1002
	rtSlidePersistAtom     = 0x03F3 // 1011
)

const recordHeaderSize = 8

// readRecordHeader reads a RecordHeader from data at the given offset.
func readRecordHeader(data []byte, offset uint32) (recordHeader, error) {
	if uint32(len(data)) < offset+recordHeaderSize {
		return recordHeader{}, errors.New("not enough data for record header")
	}
	return recordHeader{
		recVerAndInstance: binary.LittleEndian.Uint16(data[offset:]),
		recType:           binary.LittleEndian.Uint16(data[offset+2:]),
		recLen:             binary.LittleEndian.Uint32(data[offset+4:]),
	}, nil
}

// recVer returns the low 4 bits of recVerAndInstance.
func (rh recordHeader) recVer() uint16 {
	return rh.recVerAndInstance & 0x0F
}

// recInstance returns the high 12 bits of recVerAndInstance (shifted right by 4).
func (rh recordHeader) recInstance() uint16 {
	return rh.recVerAndInstance >> 4
}
