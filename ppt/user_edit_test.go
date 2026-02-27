package ppt

import (
	"encoding/binary"
	"testing"
)

// buildCurrentUserStream creates a minimal Current User stream with the given offsetToCurrentEdit.
func buildCurrentUserStream(offsetToCurrentEdit uint32) []byte {
	data := make([]byte, 20)
	// RecordHeader (8 bytes): recVerAndInstance=0, recType=rtCurrentUserAtom, recLen=12
	binary.LittleEndian.PutUint16(data[0:], 0x0000)
	binary.LittleEndian.PutUint16(data[2:], rtCurrentUserAtom)
	binary.LittleEndian.PutUint32(data[4:], 12)
	// size at offset 8 (fixed 0x14)
	binary.LittleEndian.PutUint32(data[8:], 0x14)
	// headerToken at offset 12 (unencrypted)
	binary.LittleEndian.PutUint32(data[12:], 0xE391C9F3)
	// offsetToCurrentEdit at offset 16
	binary.LittleEndian.PutUint32(data[16:], offsetToCurrentEdit)
	return data
}

func TestParseCurrentUser(t *testing.T) {
	data := buildCurrentUserStream(0x1234)
	offset, err := parseCurrentUser(data)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if offset != 0x1234 {
		t.Errorf("offsetToCurrentEdit = 0x%X, want 0x1234", offset)
	}
}

func TestParseCurrentUserTooShort(t *testing.T) {
	data := make([]byte, 10) // less than 20 bytes
	_, err := parseCurrentUser(data)
	if err == nil {
		t.Fatal("expected error for short data, got nil")
	}
}

// buildPersistDirAtom creates a PersistDirectoryAtom with the given entries.
// Each entry is (persistId, offsets...).
func buildPersistDirAtom(entries []struct {
	persistId uint32
	offsets   []uint32
}) []byte {
	// Calculate body size
	bodySize := uint32(0)
	for _, e := range entries {
		bodySize += 4 + uint32(len(e.offsets))*4 // 4 for entry header + 4*count for offsets
	}

	data := make([]byte, recordHeaderSize+bodySize)
	// RecordHeader
	binary.LittleEndian.PutUint16(data[0:], 0x0000)
	binary.LittleEndian.PutUint16(data[2:], rtPersistDirectoryAtom)
	binary.LittleEndian.PutUint32(data[4:], bodySize)

	pos := uint32(recordHeaderSize)
	for _, e := range entries {
		cPersist := uint32(len(e.offsets))
		entryHeader := (e.persistId & 0x000FFFFF) | ((cPersist & 0xFFF) << 20)
		binary.LittleEndian.PutUint32(data[pos:], entryHeader)
		pos += 4
		for _, off := range e.offsets {
			binary.LittleEndian.PutUint32(data[pos:], off)
			pos += 4
		}
	}
	return data
}

// buildUserEdit creates a UserEdit record at the start of the returned byte slice.
func buildUserEdit(offsetLastEdit, offsetPersistDir uint32) []byte {
	data := make([]byte, 24)
	// RecordHeader: recType=rtUserEdit, recLen=16
	binary.LittleEndian.PutUint16(data[0:], 0x0000)
	binary.LittleEndian.PutUint16(data[2:], rtUserEdit)
	binary.LittleEndian.PutUint32(data[4:], 16) // 24 - 8 = 16 bytes of body
	// lastSlideIdRef at offset 8
	binary.LittleEndian.PutUint32(data[8:], 0)
	// version at offset 12
	binary.LittleEndian.PutUint16(data[12:], 0)
	// minorVersion at offset 14, majorVersion at offset 15
	data[14] = 0
	data[15] = 0
	// offsetLastEdit at offset 16
	binary.LittleEndian.PutUint32(data[16:], offsetLastEdit)
	// offsetPersistDir at offset 20
	binary.LittleEndian.PutUint32(data[20:], offsetPersistDir)
	return data
}

func TestBuildPersistDirectorySingleUserEdit(t *testing.T) {
	// Build a PPT Document stream with:
	//   offset 0: UserEdit (offsetLastEdit=0, offsetPersistDir=24)
	//   offset 24: PersistDirectoryAtom with entries {persistId=0, offsets=[100, 200]}
	persistDir := buildPersistDirAtom([]struct {
		persistId uint32
		offsets   []uint32
	}{
		{persistId: 0, offsets: []uint32{100, 200}},
	})

	userEdit := buildUserEdit(0, 24) // offsetLastEdit=0 (end of chain), persistDir at offset 24

	pptDocData := make([]byte, 24+len(persistDir))
	copy(pptDocData[0:], userEdit)
	copy(pptDocData[24:], persistDir)

	result, err := buildPersistDirectory(pptDocData, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Expect persistId 0 → 100, persistId 1 → 200
	if result[0] != 100 {
		t.Errorf("persistId 0 = %d, want 100", result[0])
	}
	if result[1] != 200 {
		t.Errorf("persistId 1 = %d, want 200", result[1])
	}
	if len(result) != 2 {
		t.Errorf("len(result) = %d, want 2", len(result))
	}
}

func TestBuildPersistDirectoryChainedUserEdits(t *testing.T) {
	// Build a PPT Document stream with two UserEdits chained together.
	// UserEdit1 at offset 0: offsetLastEdit=100, offsetPersistDir=24
	// PersistDir1 at offset 24: {persistId=0, offsets=[500]}
	// UserEdit2 at offset 100: offsetLastEdit=0, offsetPersistDir=124
	// PersistDir2 at offset 124: {persistId=0, offsets=[999]}, {persistId=5, offsets=[800]}

	persistDir1 := buildPersistDirAtom([]struct {
		persistId uint32
		offsets   []uint32
	}{
		{persistId: 0, offsets: []uint32{500}},
	})

	persistDir2 := buildPersistDirAtom([]struct {
		persistId uint32
		offsets   []uint32
	}{
		{persistId: 0, offsets: []uint32{999}},
		{persistId: 5, offsets: []uint32{800}},
	})

	userEdit1 := buildUserEdit(100, 24)  // chain to UserEdit2 at offset 100
	userEdit2 := buildUserEdit(0, 124)   // end of chain, persistDir at 124

	// Total size: 24 (UE1) + len(PD1) + padding to offset 100 + 24 (UE2) + len(PD2)
	totalSize := uint32(124) + uint32(len(persistDir2))
	pptDocData := make([]byte, totalSize)
	copy(pptDocData[0:], userEdit1)
	copy(pptDocData[24:], persistDir1)
	copy(pptDocData[100:], userEdit2)
	copy(pptDocData[124:], persistDir2)

	result, err := buildPersistDirectory(pptDocData, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// First UserEdit takes precedence: persistId 0 → 500 (not 999)
	if result[0] != 500 {
		t.Errorf("persistId 0 = %d, want 500 (first UserEdit takes precedence)", result[0])
	}
	// persistId 5 only in second UserEdit
	if result[5] != 800 {
		t.Errorf("persistId 5 = %d, want 800", result[5])
	}
}

func TestBuildPersistDirectoryInvalidRecordType(t *testing.T) {
	// Build data with wrong recType at the UserEdit position
	data := make([]byte, 24)
	binary.LittleEndian.PutUint16(data[0:], 0x0000)
	binary.LittleEndian.PutUint16(data[2:], 0x0001) // wrong recType
	binary.LittleEndian.PutUint32(data[4:], 16)

	_, err := buildPersistDirectory(data, 0)
	if err == nil {
		t.Fatal("expected error for wrong recType, got nil")
	}
}

func TestBuildPersistDirectoryMultipleEntries(t *testing.T) {
	// Test PersistDirectoryAtom with multiple entries in a single atom
	persistDir := buildPersistDirAtom([]struct {
		persistId uint32
		offsets   []uint32
	}{
		{persistId: 1, offsets: []uint32{100, 200, 300}},
		{persistId: 10, offsets: []uint32{400, 500}},
	})

	userEdit := buildUserEdit(0, 24)

	pptDocData := make([]byte, 24+len(persistDir))
	copy(pptDocData[0:], userEdit)
	copy(pptDocData[24:], persistDir)

	result, err := buildPersistDirectory(pptDocData, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	expected := map[uint32]uint32{
		1:  100,
		2:  200,
		3:  300,
		10: 400,
		11: 500,
	}
	if len(result) != len(expected) {
		t.Errorf("len(result) = %d, want %d", len(result), len(expected))
	}
	for id, off := range expected {
		if result[id] != off {
			t.Errorf("persistId %d = %d, want %d", id, result[id], off)
		}
	}
}
