package ppt

import (
	"encoding/binary"
	"testing"
)

func TestReadRecordHeader(t *testing.T) {
	// Build an 8-byte record header: recVerAndInstance=0x0021, recType=0x0FF0, recLen=100
	data := make([]byte, 16)
	binary.LittleEndian.PutUint16(data[4:], 0x0021)  // recVerAndInstance at offset 4
	binary.LittleEndian.PutUint16(data[6:], 0x0FF0)  // recType
	binary.LittleEndian.PutUint32(data[8:], 100)      // recLen

	rh, err := readRecordHeader(data, 4)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if rh.recType != rtSlideListWithText {
		t.Errorf("recType = 0x%04X, want 0x%04X", rh.recType, rtSlideListWithText)
	}
	if rh.recLen != 100 {
		t.Errorf("recLen = %d, want 100", rh.recLen)
	}
	// recVerAndInstance = 0x0021 → recVer = 0x1, recInstance = 0x002
	if rh.recVer() != 0x01 {
		t.Errorf("recVer() = %d, want 1", rh.recVer())
	}
	if rh.recInstance() != 0x02 {
		t.Errorf("recInstance() = %d, want 2", rh.recInstance())
	}
}

func TestReadRecordHeaderNotEnoughData(t *testing.T) {
	data := make([]byte, 6) // less than 8 bytes
	_, err := readRecordHeader(data, 0)
	if err == nil {
		t.Fatal("expected error for insufficient data, got nil")
	}
}

func TestReadRecordHeaderOffsetOverflow(t *testing.T) {
	data := make([]byte, 10)
	_, err := readRecordHeader(data, 5) // 5+8=13 > 10
	if err == nil {
		t.Fatal("expected error for offset overflow, got nil")
	}
}
