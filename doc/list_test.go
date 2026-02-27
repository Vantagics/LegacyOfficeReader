package doc

import (
	"encoding/binary"
	"math/rand"
	"testing"
	"testing/quick"
)

// buildPlcfLst constructs a PlcfLst binary blob from list definitions.
// It builds the LSTF array and a minimal LVLF per list (fSimpleList=1)
// so that parsePlcfLst can determine ordered/unordered from nfc.
func buildPlcfLst(lists []listDef) []byte {
	cLst := len(lists)
	// Header: 2 bytes (cLst)
	// LSTF entries: cLst * 28 bytes each
	// LVLF entries: 1 per list (fSimpleList=1), each 28 bytes LVLF + 2 bytes xst length (0)
	lstfSize := cLst * 28
	lvlfSize := cLst * (28 + 2) // 28 bytes LVLF + 2 bytes xst (length=0)
	totalSize := 2 + lstfSize + lvlfSize
	data := make([]byte, totalSize)

	// Write cLst
	binary.LittleEndian.PutUint16(data[0:2], uint16(cLst))

	// Write LSTF entries
	for i, ld := range lists {
		off := 2 + i*28
		binary.LittleEndian.PutUint32(data[off:off+4], ld.listID)
		// offset 26: flags byte - bit 0 = fSimpleList (set to 1 so only 1 LVLF)
		data[off+26] = 0x01
	}

	// Write LVLF entries after LSTF array
	lvlfPos := 2 + lstfSize
	for _, ld := range lists {
		// LVLF: 28 bytes
		// offset 0: iStartAt (uint32) = 1
		binary.LittleEndian.PutUint32(data[lvlfPos:lvlfPos+4], 1)
		// offset 4: nfc (uint8) - 23=bullet(unordered), 0=decimal(ordered)
		if ld.ordered {
			data[lvlfPos+4] = 0 // decimal = ordered
		} else {
			data[lvlfPos+4] = 23 // bullet = unordered
		}
		// offset 24: cbGrpprlChpx = 0
		data[lvlfPos+24] = 0
		// offset 25: cbGrpprlPapx = 0
		data[lvlfPos+25] = 0
		lvlfPos += 28
		// xst: uint16 length = 0 (no number text)
		binary.LittleEndian.PutUint16(data[lvlfPos:lvlfPos+2], 0)
		lvlfPos += 2
	}

	return data
}

// buildPlfLfo constructs a PlfLfo binary blob from list overrides.
func buildPlfLfo(overrides []listOverride) []byte {
	lfoMac := len(overrides)
	// Header: 4 bytes (lfoMac)
	// LFO entries: lfoMac * 16 bytes each
	totalSize := 4 + lfoMac*16
	data := make([]byte, totalSize)

	binary.LittleEndian.PutUint32(data[0:4], uint32(lfoMac))

	for i, lo := range overrides {
		off := 4 + i*16
		binary.LittleEndian.PutUint32(data[off:off+4], lo.listID)
		// Remaining 12 bytes are zero (other LFO fields, ignored)
	}

	return data
}

// **Feature: doc-format-preservation, Property 6: 列表识别**
// **Validates: Requirements 7.1, 7.2, 7.3, 7.4, 7.5**
func TestPropertyListIdentification(t *testing.T) {
	cfg := &quick.Config{
		MaxCount: 100,
		Rand:     rand.New(rand.NewSource(42)),
	}

	f := func(seed int64) bool {
		r := rand.New(rand.NewSource(seed))

		// Generate 1-5 random list definitions
		numLists := r.Intn(5) + 1
		expectedLists := make([]listDef, numLists)
		for i := 0; i < numLists; i++ {
			expectedLists[i] = listDef{
				listID:  r.Uint32(),
				ordered: r.Intn(2) == 1,
			}
		}

		// Build PlcfLst binary data and parse it
		plcfData := buildPlcfLst(expectedLists)
		parsedLists, err := parsePlcfLst(plcfData, 0, uint32(len(plcfData)))
		if err != nil {
			t.Logf("parsePlcfLst error: %v", err)
			return false
		}
		if len(parsedLists) != numLists {
			t.Logf("parsePlcfLst: got %d lists, want %d", len(parsedLists), numLists)
			return false
		}
		for i := 0; i < numLists; i++ {
			if parsedLists[i].listID != expectedLists[i].listID {
				t.Logf("list[%d].listID: got %d, want %d", i, parsedLists[i].listID, expectedLists[i].listID)
				return false
			}
			if parsedLists[i].ordered != expectedLists[i].ordered {
				t.Logf("list[%d].ordered: got %v, want %v", i, parsedLists[i].ordered, expectedLists[i].ordered)
				return false
			}
		}

		// Generate 1-5 random list overrides referencing the list definitions
		numOverrides := r.Intn(5) + 1
		expectedOverrides := make([]listOverride, numOverrides)
		for i := 0; i < numOverrides; i++ {
			// Reference a random list's ID
			idx := r.Intn(numLists)
			expectedOverrides[i] = listOverride{
				listID: expectedLists[idx].listID,
			}
		}

		// Build PlfLfo binary data and parse it
		lfoData := buildPlfLfo(expectedOverrides)
		parsedOverrides, err := parsePlfLfo(lfoData, 0, uint32(len(lfoData)))
		if err != nil {
			t.Logf("parsePlfLfo error: %v", err)
			return false
		}
		if len(parsedOverrides) != numOverrides {
			t.Logf("parsePlfLfo: got %d overrides, want %d", len(parsedOverrides), numOverrides)
			return false
		}
		for i := 0; i < numOverrides; i++ {
			if parsedOverrides[i].listID != expectedOverrides[i].listID {
				t.Logf("override[%d].listID: got %d, want %d", i, parsedOverrides[i].listID, expectedOverrides[i].listID)
				return false
			}
		}

		// Test empty case: lcb=0 should return empty slices
		emptyLists, err := parsePlcfLst(plcfData, 0, 0)
		if err != nil {
			t.Logf("parsePlcfLst(lcb=0) error: %v", err)
			return false
		}
		if len(emptyLists) != 0 {
			t.Logf("parsePlcfLst(lcb=0): got %d lists, want 0", len(emptyLists))
			return false
		}

		emptyOverrides, err := parsePlfLfo(lfoData, 0, 0)
		if err != nil {
			t.Logf("parsePlfLfo(lcb=0) error: %v", err)
			return false
		}
		if len(emptyOverrides) != 0 {
			t.Logf("parsePlfLfo(lcb=0): got %d overrides, want 0", len(emptyOverrides))
			return false
		}

		return true
	}

	if err := quick.Check(f, cfg); err != nil {
		t.Errorf("Property test failed: %v", err)
	}
}

func TestParsePlcfLst_ValidData(t *testing.T) {
	lists := []listDef{
		{listID: 100, ordered: true},
		{listID: 200, ordered: false},
		{listID: 300, ordered: true},
	}
	data := buildPlcfLst(lists)

	result, err := parsePlcfLst(data, 0, uint32(len(data)))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(result) != 3 {
		t.Fatalf("got %d lists, want 3", len(result))
	}
	for i, want := range lists {
		if result[i].listID != want.listID {
			t.Errorf("list[%d].listID = %d, want %d", i, result[i].listID, want.listID)
		}
		if result[i].ordered != want.ordered {
			t.Errorf("list[%d].ordered = %v, want %v", i, result[i].ordered, want.ordered)
		}
	}
}

func TestParsePlfLfo_ValidData(t *testing.T) {
	overrides := []listOverride{
		{listID: 100},
		{listID: 200},
		{listID: 300},
	}
	data := buildPlfLfo(overrides)

	result, err := parsePlfLfo(data, 0, uint32(len(data)))
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(result) != 3 {
		t.Fatalf("got %d overrides, want 3", len(result))
	}
	for i, want := range overrides {
		if result[i].listID != want.listID {
			t.Errorf("override[%d].listID = %d, want %d", i, result[i].listID, want.listID)
		}
	}
}

func TestParsePlcfLst_Empty(t *testing.T) {
	tableData := make([]byte, 100)
	result, err := parsePlcfLst(tableData, 0, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(result) != 0 {
		t.Errorf("got %d lists, want 0", len(result))
	}
}

func TestParsePlfLfo_Empty(t *testing.T) {
	tableData := make([]byte, 100)
	result, err := parsePlfLfo(tableData, 0, 0)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if len(result) != 0 {
		t.Errorf("got %d overrides, want 0", len(result))
	}
}
