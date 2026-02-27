package doc

import (
	"encoding/binary"
	"math/rand"
	"testing"
	"testing/quick"
)

func TestParseFIB_ValidSignature(t *testing.T) {
	// Build a minimal valid FIB buffer.
	// FIB base (32 bytes) + 2 (csw) + csw*2 (FibRgW) + 2 (cslw) + cslw*4 (FibRgLw) + 2 (cbRgFcLcb) + cbRgFcLcb*4 (FibRgFcLcb uint32 values)
	// cbRgFcLcb is the count of uint32 values. We need at least 68 (indices 66 and 67 for fcClx/lcbClx).
	csw := uint16(14)
	cslw := uint16(22)
	cbRgFcLcb := uint16(93) // Word 97 standard: 93 uint32 values

	totalSize := 32 + 2 + int(csw)*2 + 2 + int(cslw)*4 + 2 + int(cbRgFcLcb)*4
	data := make([]byte, totalSize)

	// wIdent at offset 0x00
	binary.LittleEndian.PutUint16(data[0x00:], 0xA5EC)
	// nFib at offset 0x02
	binary.LittleEndian.PutUint16(data[0x02:], 0x00C1)
	// flags at offset 0x0A: set bit 9 (fWhichTblStm = 1)
	binary.LittleEndian.PutUint16(data[0x0A:], 0x0200)

	// csw at offset 0x20
	offset := 0x20
	binary.LittleEndian.PutUint16(data[offset:], csw)
	offset += 2 + int(csw)*2

	// cslw
	binary.LittleEndian.PutUint16(data[offset:], cslw)
	offset += 2 + int(cslw)*4

	// cbRgFcLcb
	binary.LittleEndian.PutUint16(data[offset:], cbRgFcLcb)
	offset += 2

	// fcClx at uint32 index 66, lcbClx at uint32 index 67
	fcClxOffset := offset + 66*4
	binary.LittleEndian.PutUint32(data[fcClxOffset:], 0x1234)
	binary.LittleEndian.PutUint32(data[fcClxOffset+4:], 0x5678)

	f, err := parseFIB(data)
	if err != nil {
		t.Fatalf("parseFIB returned unexpected error: %v", err)
	}
	if f.wIdent != 0xA5EC {
		t.Errorf("wIdent = 0x%04X, want 0xA5EC", f.wIdent)
	}
	if f.nFib != 0x00C1 {
		t.Errorf("nFib = 0x%04X, want 0x00C1", f.nFib)
	}
	if f.fWhichTblStm != 1 {
		t.Errorf("fWhichTblStm = %d, want 1", f.fWhichTblStm)
	}
	if f.fcClx != 0x1234 {
		t.Errorf("fcClx = 0x%08X, want 0x00001234", f.fcClx)
	}
	if f.lcbClx != 0x5678 {
		t.Errorf("lcbClx = 0x%08X, want 0x00005678", f.lcbClx)
	}
}

func TestParseFIB_Table0(t *testing.T) {
	// Same as above but with fWhichTblStm = 0 (bit 9 not set)
	csw := uint16(14)
	cslw := uint16(22)
	cbRgFcLcb := uint16(93)

	totalSize := 32 + 2 + int(csw)*2 + 2 + int(cslw)*4 + 2 + int(cbRgFcLcb)*4
	data := make([]byte, totalSize)

	binary.LittleEndian.PutUint16(data[0x00:], 0xA5EC)
	binary.LittleEndian.PutUint16(data[0x02:], 0x00C1)
	// flags at 0x0A: bit 9 NOT set
	binary.LittleEndian.PutUint16(data[0x0A:], 0x0000)

	offset := 0x20
	binary.LittleEndian.PutUint16(data[offset:], csw)
	offset += 2 + int(csw)*2
	binary.LittleEndian.PutUint16(data[offset:], cslw)
	offset += 2 + int(cslw)*4
	binary.LittleEndian.PutUint16(data[offset:], cbRgFcLcb)

	f, err := parseFIB(data)
	if err != nil {
		t.Fatalf("parseFIB returned unexpected error: %v", err)
	}
	if f.fWhichTblStm != 0 {
		t.Errorf("fWhichTblStm = %d, want 0", f.fWhichTblStm)
	}
}

func TestParseFIB_InvalidSignature(t *testing.T) {
	data := make([]byte, 1024)
	binary.LittleEndian.PutUint16(data[0x00:], 0xBEEF) // wrong signature

	_, err := parseFIB(data)
	if err == nil {
		t.Fatal("parseFIB should return error for invalid signature")
	}
}

func TestParseFIB_TooShort(t *testing.T) {
	data := make([]byte, 10) // way too short
	_, err := parseFIB(data)
	if err == nil {
		t.Fatal("parseFIB should return error for data too short")
	}
}

// buildFIBBuffer creates a valid FIB buffer with the given cbRgFcLcb count
// and writes format field values at the correct uint32 indices within FibRgFcLcb.
// lwFields optionally writes uint32 values at indices within FibRgLw97.
// Returns the buffer and the byte offset where FibRgFcLcb starts.
func buildFIBBuffer(cbRgFcLcb uint16, formatFields map[int]uint32) []byte {
	return buildFIBBufferFull(cbRgFcLcb, formatFields, nil)
}

func buildFIBBufferFull(cbRgFcLcb uint16, formatFields map[int]uint32, lwFields map[int]uint32) []byte {
	csw := uint16(14)
	cslw := uint16(22)

	totalSize := 32 + 2 + int(csw)*2 + 2 + int(cslw)*4 + 2 + int(cbRgFcLcb)*4
	data := make([]byte, totalSize)

	// FIB base
	binary.LittleEndian.PutUint16(data[0x00:], 0xA5EC) // wIdent
	binary.LittleEndian.PutUint16(data[0x02:], 0x00C1) // nFib
	binary.LittleEndian.PutUint16(data[0x0A:], 0x0200) // flags: fWhichTblStm=1

	offset := 0x20
	binary.LittleEndian.PutUint16(data[offset:], csw)
	offset += 2 + int(csw)*2

	binary.LittleEndian.PutUint16(data[offset:], cslw)
	lwStart := offset + 2

	// Write FibRgLw97 fields
	for idx, val := range lwFields {
		if idx < int(cslw) {
			off := lwStart + idx*4
			binary.LittleEndian.PutUint32(data[off:], val)
		}
	}

	offset += 2 + int(cslw)*4

	binary.LittleEndian.PutUint16(data[offset:], cbRgFcLcb)
	offset += 2

	// Write format fields at specified indices
	for idx, val := range formatFields {
		if idx < int(cbRgFcLcb) {
			off := offset + idx*4
			binary.LittleEndian.PutUint32(data[off:], val)
		}
	}

	return data
}

// **Feature: doc-format-preservation, Property 1: FIB 格式字段提取**
// **Validates: Requirements 1.1, 1.2, 1.3, 1.4, 1.5**
// FibRgFcLcb97 uint32 indices per [MS-DOC]:
//   fcStshf=2, lcbStshf=3, fcPlcfBteChpx=24, lcbPlcfBteChpx=25,
//   fcPlcfBtePapx=26, lcbPlcfBtePapx=27, fcSttbfFfn=30, lcbSttbfFfn=31,
//   fcPlcfHdd=22, lcbPlcfHdd=23, fcPlfLst=146, lcbPlfLst=147,
//   fcPlfLfo=148, lcbPlfLfo=149, fcClx=66, lcbClx=67
// FibRgLw97 uint32 indices: ccpText=3, ccpFtn=4, ccpHdd=5
func TestPropertyFIB_FormatFieldExtraction(t *testing.T) {
	cfg := &quick.Config{
		MaxCount: 100,
		Rand:     rand.New(rand.NewSource(42)),
	}

	f := func(
		fcStshf, lcbStshf,
		fcPlcfBteChpx, lcbPlcfBteChpx,
		fcPlcfBtePapx, lcbPlcfBtePapx,
		fcSttbfFfn, lcbSttbfFfn,
		fcPlcfHdd, lcbPlcfHdd,
		fcPlcfLst, lcbPlcfLst,
		fcPlfLfo, lcbPlfLfo,
		fcClx, lcbClx,
		ccpText, ccpFtn, ccpHdd uint32,
	) bool {
		fields := map[int]uint32{
			2:   fcStshf,
			3:   lcbStshf,
			22:  fcPlcfHdd,
			23:  lcbPlcfHdd,
			24:  fcPlcfBteChpx,
			25:  lcbPlcfBteChpx,
			26:  fcPlcfBtePapx,
			27:  lcbPlcfBtePapx,
			30:  fcSttbfFfn,
			31:  lcbSttbfFfn,
			146: fcPlcfLst,
			147: lcbPlcfLst,
			148: fcPlfLfo,
			149: lcbPlfLfo,
			66:  fcClx,
			67:  lcbClx,
		}

		lwFields := map[int]uint32{
			3: ccpText,
			4: ccpFtn,
			5: ccpHdd,
		}

		// cbRgFcLcb must be >= 150 to cover fcPlfLfo at index 149
		data := buildFIBBufferFull(186, fields, lwFields)
		result, err := parseFIB(data)
		if err != nil {
			t.Logf("parseFIB error: %v", err)
			return false
		}

		return result.fcStshf == fcStshf &&
			result.lcbStshf == lcbStshf &&
			result.fcPlcfHdd == fcPlcfHdd &&
			result.lcbPlcfHdd == lcbPlcfHdd &&
			result.fcPlcfBteChpx == fcPlcfBteChpx &&
			result.lcbPlcfBteChpx == lcbPlcfBteChpx &&
			result.fcPlcfBtePapx == fcPlcfBtePapx &&
			result.lcbPlcfBtePapx == lcbPlcfBtePapx &&
			result.fcSttbfFfn == fcSttbfFfn &&
			result.lcbSttbfFfn == lcbSttbfFfn &&
			result.fcPlcfLst == fcPlcfLst &&
			result.lcbPlcfLst == lcbPlcfLst &&
			result.fcPlfLfo == fcPlfLfo &&
			result.lcbPlfLfo == lcbPlfLfo &&
			result.fcClx == fcClx &&
			result.lcbClx == lcbClx &&
			result.ccpText == ccpText &&
			result.ccpFtn == ccpFtn &&
			result.ccpHdd == ccpHdd
	}

	if err := quick.Check(f, cfg); err != nil {
		t.Errorf("Property test failed: %v", err)
	}
}

func TestParseFIB_FormatFields(t *testing.T) {
	fields := map[int]uint32{
		2:   0xAAAA0001, // fcStshf
		3:   0xAAAA0002, // lcbStshf
		22:  0xFFFF0001, // fcPlcfHdd
		23:  0xFFFF0002, // lcbPlcfHdd
		24:  0xBBBB0001, // fcPlcfBteChpx
		25:  0xBBBB0002, // lcbPlcfBteChpx
		26:  0xCCCC0001, // fcPlcfBtePapx
		27:  0xCCCC0002, // lcbPlcfBtePapx
		30:  0x11110001, // fcSttbfFfn
		31:  0x11110002, // lcbSttbfFfn
		146: 0xDDDD0001, // fcPlfLst
		147: 0xDDDD0002, // lcbPlfLst
		148: 0xEEEE0001, // fcPlfLfo
		149: 0xEEEE0002, // lcbPlfLfo
		66:  0x00001234, // fcClx
		67:  0x00005678, // lcbClx
	}

	lwFields := map[int]uint32{
		3: 0x00000100, // ccpText
		4: 0x00000050, // ccpFtn
		5: 0x00000080, // ccpHdd
	}

	data := buildFIBBufferFull(186, fields, lwFields)
	f, err := parseFIB(data)
	if err != nil {
		t.Fatalf("parseFIB returned unexpected error: %v", err)
	}

	checks := []struct {
		name string
		got  uint32
		want uint32
	}{
		{"fcStshf", f.fcStshf, 0xAAAA0001},
		{"lcbStshf", f.lcbStshf, 0xAAAA0002},
		{"fcPlcfHdd", f.fcPlcfHdd, 0xFFFF0001},
		{"lcbPlcfHdd", f.lcbPlcfHdd, 0xFFFF0002},
		{"fcPlcfBteChpx", f.fcPlcfBteChpx, 0xBBBB0001},
		{"lcbPlcfBteChpx", f.lcbPlcfBteChpx, 0xBBBB0002},
		{"fcPlcfBtePapx", f.fcPlcfBtePapx, 0xCCCC0001},
		{"lcbPlcfBtePapx", f.lcbPlcfBtePapx, 0xCCCC0002},
		{"fcSttbfFfn", f.fcSttbfFfn, 0x11110001},
		{"lcbSttbfFfn", f.lcbSttbfFfn, 0x11110002},
		{"fcPlcfLst", f.fcPlcfLst, 0xDDDD0001},
		{"lcbPlcfLst", f.lcbPlcfLst, 0xDDDD0002},
		{"fcPlfLfo", f.fcPlfLfo, 0xEEEE0001},
		{"lcbPlfLfo", f.lcbPlfLfo, 0xEEEE0002},
		{"ccpText", f.ccpText, 0x00000100},
		{"ccpFtn", f.ccpFtn, 0x00000050},
		{"ccpHdd", f.ccpHdd, 0x00000080},
	}

	for _, c := range checks {
		if c.got != c.want {
			t.Errorf("%s = 0x%08X, want 0x%08X", c.name, c.got, c.want)
		}
	}
}

func TestParseFIB_SmallFibRgFcLcb(t *testing.T) {
	// cbRgFcLcb = 68 is just enough for fcClx/lcbClx (indices 66, 67)
	// but too small for fcPlfLst (index 146) and fcPlfLfo (index 148).
	// Format fields at indices 2-27 (STSH, CHPX, PAPX) are within range.
	fields := map[int]uint32{
		66: 0x1234, // fcClx
		67: 0x5678, // lcbClx
	}

	data := buildFIBBuffer(68, fields)
	f, err := parseFIB(data)
	if err != nil {
		t.Fatalf("parseFIB returned unexpected error: %v", err)
	}

	// Format fields at indices 2-27 should be zero since we didn't write them
	if f.fcStshf != 0 {
		t.Errorf("fcStshf = 0x%08X, want 0", f.fcStshf)
	}
	if f.lcbStshf != 0 {
		t.Errorf("lcbStshf = 0x%08X, want 0", f.lcbStshf)
	}
	if f.fcPlcfBteChpx != 0 {
		t.Errorf("fcPlcfBteChpx = 0x%08X, want 0", f.fcPlcfBteChpx)
	}
	if f.lcbPlcfBteChpx != 0 {
		t.Errorf("lcbPlcfBteChpx = 0x%08X, want 0", f.lcbPlcfBteChpx)
	}
	if f.fcPlcfBtePapx != 0 {
		t.Errorf("fcPlcfBtePapx = 0x%08X, want 0", f.fcPlcfBtePapx)
	}
	if f.lcbPlcfBtePapx != 0 {
		t.Errorf("lcbPlcfBtePapx = 0x%08X, want 0", f.lcbPlcfBtePapx)
	}
	// List fields at indices 146-149 are beyond cbRgFcLcb=68, so should be zero
	if f.fcPlcfLst != 0 {
		t.Errorf("fcPlcfLst = 0x%08X, want 0", f.fcPlcfLst)
	}
	if f.lcbPlcfLst != 0 {
		t.Errorf("lcbPlcfLst = 0x%08X, want 0", f.lcbPlcfLst)
	}
	if f.fcPlfLfo != 0 {
		t.Errorf("fcPlfLfo = 0x%08X, want 0", f.fcPlfLfo)
	}
	if f.lcbPlfLfo != 0 {
		t.Errorf("lcbPlfLfo = 0x%08X, want 0", f.lcbPlfLfo)
	}

	// fcClx/lcbClx should be correctly read
	if f.fcClx != 0x1234 {
		t.Errorf("fcClx = 0x%08X, want 0x00001234", f.fcClx)
	}
	if f.lcbClx != 0x5678 {
		t.Errorf("lcbClx = 0x%08X, want 0x00005678", f.lcbClx)
	}
}
