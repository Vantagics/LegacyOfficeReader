package doc

import (
	"encoding/binary"
	"errors"
	"fmt"
)

const fibSignature = 0xA5EC

// fib represents the File Information Block of a DOC file.
type fib struct {
	wIdent       uint16
	nFib         uint16
	fWhichTblStm uint8
	lid          uint16 // language ID
	fcClx        uint32
	lcbClx       uint32
	// Character counts from FibRgLw97
	ccpText uint32 // character count of main document text
	ccpFtn  uint32 // character count of footnote text
	ccpHdd  uint32 // character count of header/footer text
	// Format-related fields from FibRgFcLcb
	fcStshf        uint32
	lcbStshf       uint32
	fcPlcfHdd      uint32 // header/footer text positions
	lcbPlcfHdd     uint32
	fcPlcfBteChpx  uint32
	lcbPlcfBteChpx uint32
	fcPlcfBtePapx  uint32
	lcbPlcfBtePapx uint32
	fcSttbfFfn     uint32 // font name table
	lcbSttbfFfn    uint32
	fcPlcfLst      uint32
	lcbPlcfLst     uint32
	fcPlfLfo       uint32
	lcbPlfLfo      uint32
	fcPlcfSed      uint32 // section descriptor positions
	lcbPlcfSed     uint32
	fcDggInfo      uint32 // OfficeArt Drawing Group info
	lcbDggInfo     uint32
	fcPlcSpaMom    uint32 // main document shape positions
	lcbPlcSpaMom   uint32
	fcPlcSpaHdr    uint32 // header/footer shape positions
	lcbPlcSpaHdr   uint32
	fcPlcftxbxTxt  uint32 // text box text positions
	lcbPlcftxbxTxt uint32
}

// parseFIB parses the FIB from the WordDocument stream data.
// It reads the FIB base, navigates through FibRgW and FibRgLw,
// and locates fcClx/lcbClx in FibRgFcLcb.
func parseFIB(wordDocData []byte) (fib, error) {
	// FIB base is 32 bytes; we need at least that plus the csw field
	if len(wordDocData) < 34 {
		return fib{}, errors.New("WordDocument stream too short for FIB")
	}

	// Read wIdent at offset 0x00
	wIdent := binary.LittleEndian.Uint16(wordDocData[0x00:])
	if wIdent != fibSignature {
		return fib{}, fmt.Errorf("invalid FIB signature: expected 0x%04X, got 0x%04X", fibSignature, wIdent)
	}

	// Read nFib at offset 0x02
	nFib := binary.LittleEndian.Uint16(wordDocData[0x02:])

	// Read flags at offset 0x0A, extract fWhichTblStm (bit 9)
	flags := binary.LittleEndian.Uint16(wordDocData[0x0A:])
	var fWhichTblStm uint8
	if flags&0x0200 != 0 {
		fWhichTblStm = 1
	}

	// Read lid (language ID) at offset 0x06
	var lid uint16
	if len(wordDocData) >= 0x08 {
		lid = binary.LittleEndian.Uint16(wordDocData[0x06:])
	}

	// Navigate to FibRgFcLcb:
	// FIB base is 32 bytes (0x20)
	offset := 0x20

	// Read csw (uint16) = count of uint16s in FibRgW
	if offset+2 > len(wordDocData) {
		return fib{}, errors.New("WordDocument stream too short for FibRgW csw")
	}
	csw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2

	// Skip FibRgW: csw * 2 bytes
	offset += int(csw) * 2
	if offset+2 > len(wordDocData) {
		return fib{}, errors.New("WordDocument stream too short for FibRgLw cslw")
	}

	// Read cslw (uint16) = count of uint32s in FibRgLw
	cslw := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2

	// Read FibRgLw97 fields before skipping
	fibRgLwStart := offset
	var ccpText, ccpFtn, ccpHdd uint32
	// FibRgLw97 per [MS-DOC]: index 3 = ccpText, index 4 = ccpFtn, index 5 = ccpHdd
	if int(cslw) > 3 && fibRgLwStart+4*4 <= len(wordDocData) {
		ccpText = binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+3*4:])
	}
	if int(cslw) > 4 && fibRgLwStart+5*4 <= len(wordDocData) {
		ccpFtn = binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+4*4:])
	}
	if int(cslw) > 5 && fibRgLwStart+6*4 <= len(wordDocData) {
		ccpHdd = binary.LittleEndian.Uint32(wordDocData[fibRgLwStart+5*4:])
	}

	// Skip FibRgLw: cslw * 4 bytes
	offset += int(cslw) * 4
	if offset+2 > len(wordDocData) {
		return fib{}, errors.New("WordDocument stream too short for FibRgFcLcb cbRgFcLcb")
	}

	// Read cbRgFcLcb (uint16) = count of FC/LCB pairs in FibRgFcLcb
	cbRgFcLcb := binary.LittleEndian.Uint16(wordDocData[offset:])
	offset += 2

	// fcClx is at uint32 index 66 in FibRgFcLcb, lcbClx at index 67.
	// cbRgFcLcb is the count of uint32 values in the FibRgFcLcb blob.
	const fcClxIndex = 66
	const lcbClxIndex = 67
	if int(cbRgFcLcb) <= lcbClxIndex {
		return fib{}, fmt.Errorf("FibRgFcLcb too small: %d uint32 values, need at least %d", cbRgFcLcb, lcbClxIndex+1)
	}

	// Helper to read a uint32 from FibRgFcLcb at a given index, if the index is within bounds.
	readFcLcb := func(index int) uint32 {
		if int(cbRgFcLcb) <= index {
			return 0
		}
		off := offset + index*4
		if off+4 > len(wordDocData) {
			return 0
		}
		return binary.LittleEndian.Uint32(wordDocData[off:])
	}

	// Extract format-related fields from FibRgFcLcb.
	// Indices are uint32 positions within the FibRgFcLcb blob, per [MS-DOC] FibRgFcLcb97.
	// If FibRgFcLcb is too small for a given field, it defaults to zero.
	fcStshf := readFcLcb(2)        // index 2 in FibRgFcLcb97
	lcbStshf := readFcLcb(3)       // index 3
	fcPlcfHdd := readFcLcb(22)     // index 22 - header/footer positions
	lcbPlcfHdd := readFcLcb(23)    // index 23
	fcPlcfBteChpx := readFcLcb(24) // index 24
	lcbPlcfBteChpx := readFcLcb(25)
	fcPlcfBtePapx := readFcLcb(26) // index 26
	lcbPlcfBtePapx := readFcLcb(27)
	fcSttbfFfn := readFcLcb(30)  // index 30 - font name table
	lcbSttbfFfn := readFcLcb(31) // index 31
	fcPlcfLst := readFcLcb(146)  // fcPlfLst, index 146
	lcbPlcfLst := readFcLcb(147)
	fcPlfLfo := readFcLcb(148) // index 148
	lcbPlfLfo := readFcLcb(149)
	fcPlcfSed := readFcLcb(12) // index 12 - section descriptor positions
	lcbPlcfSed := readFcLcb(13)
	fcDggInfo := readFcLcb(100) // index 100 - OfficeArt Drawing Group info
	lcbDggInfo := readFcLcb(101)
	fcPlcSpaMom := readFcLcb(80) // index 80 - main document shape positions
	lcbPlcSpaMom := readFcLcb(81)
	fcPlcSpaHdr := readFcLcb(82) // index 82 - header/footer shape positions
	lcbPlcSpaHdr := readFcLcb(83)
	fcPlcftxbxTxt := readFcLcb(112) // index 112 - text box text positions
	lcbPlcftxbxTxt := readFcLcb(113)

	fcClxOffset := offset + fcClxIndex*4
	if fcClxOffset+8 > len(wordDocData) {
		return fib{}, errors.New("WordDocument stream too short for fcClx/lcbClx")
	}

	fcClx := binary.LittleEndian.Uint32(wordDocData[fcClxOffset:])
	lcbClx := binary.LittleEndian.Uint32(wordDocData[fcClxOffset+4:])

	return fib{
		wIdent:         wIdent,
		nFib:           nFib,
		fWhichTblStm:   fWhichTblStm,
		lid:            lid,
		fcClx:          fcClx,
		lcbClx:         lcbClx,
		ccpText:        ccpText,
		ccpFtn:         ccpFtn,
		ccpHdd:         ccpHdd,
		fcStshf:        fcStshf,
		lcbStshf:       lcbStshf,
		fcPlcfHdd:      fcPlcfHdd,
		lcbPlcfHdd:     lcbPlcfHdd,
		fcPlcfBteChpx:  fcPlcfBteChpx,
		lcbPlcfBteChpx: lcbPlcfBteChpx,
		fcPlcfBtePapx:  fcPlcfBtePapx,
		lcbPlcfBtePapx: lcbPlcfBtePapx,
		fcSttbfFfn:     fcSttbfFfn,
		lcbSttbfFfn:    lcbSttbfFfn,
		fcPlcfLst:      fcPlcfLst,
		lcbPlcfLst:     lcbPlcfLst,
		fcPlfLfo:       fcPlfLfo,
		lcbPlfLfo:      lcbPlfLfo,
		fcPlcfSed:      fcPlcfSed,
		lcbPlcfSed:     lcbPlcfSed,
		fcDggInfo:      fcDggInfo,
		lcbDggInfo:     lcbDggInfo,
		fcPlcSpaMom:    fcPlcSpaMom,
		lcbPlcSpaMom:   lcbPlcSpaMom,
		fcPlcSpaHdr:    fcPlcSpaHdr,
		lcbPlcSpaHdr:   lcbPlcSpaHdr,
		fcPlcftxbxTxt:  fcPlcftxbxTxt,
		lcbPlcftxbxTxt: lcbPlcftxbxTxt,
	}, nil
}
