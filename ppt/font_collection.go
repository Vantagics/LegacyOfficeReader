package ppt

import (
	"encoding/binary"
	"unicode/utf16"
)

// PPT record types for font collection
const (
	rtFontCollection = 0x07D5 // 2005
	rtFontEntityAtom = 0x0FB7 // 4023
)

// parseFontCollection scans pptDocData for FontEntityAtom records inside
// a FontCollection container and returns font names in order.
func parseFontCollection(pptDocData []byte) []string {
	var fonts []string
	dataLen := uint32(len(pptDocData))
	offset := uint32(0)

	for offset+recordHeaderSize <= dataLen {
		rh, err := readRecordHeader(pptDocData, offset)
		if err != nil {
			break
		}
		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > dataLen {
			break
		}

		if rh.recType == rtFontCollection && rh.recVer() == 0xF {
			// Scan inside the FontCollection container for FontEntityAtom records
			fonts = parseFontEntityAtoms(pptDocData, recDataStart, recDataEnd)
			return fonts
		}

		if rh.recVer() == 0xF {
			offset = recDataStart // step into container
		} else {
			offset = recDataEnd // skip atom
		}
	}

	return []string{}
}

// parseFontEntityAtoms extracts font names from FontEntityAtom records
// within a FontCollection container.
func parseFontEntityAtoms(data []byte, start, end uint32) []string {
	var fonts []string
	offset := start

	for offset+recordHeaderSize <= end {
		rh, err := readRecordHeader(data, offset)
		if err != nil {
			break
		}
		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen
		if recDataEnd > end {
			break
		}

		if rh.recType == rtFontEntityAtom {
			name := parseFontEntityName(data[recDataStart:recDataEnd])
			fonts = append(fonts, name)
		}

		if rh.recVer() == 0xF {
			// FontEmbedDataBlob or other containers - step into
			offset = recDataStart
		} else {
			offset = recDataEnd
		}
	}

	if fonts == nil {
		return []string{}
	}
	return fonts
}

// parseFontEntityName extracts the lfFaceName from a FontEntityAtom body.
// The lfFaceName is a UTF-16LE string, up to 32 characters (64 bytes), null-terminated.
func parseFontEntityName(data []byte) string {
	// lfFaceName occupies the first 64 bytes of FontEntityAtom
	nameLen := 64
	if len(data) < nameLen {
		nameLen = len(data)
	}
	nameData := data[:nameLen]

	// Decode UTF-16LE, stopping at null terminator
	numChars := nameLen / 2
	chars := make([]uint16, numChars)
	for i := 0; i < numChars; i++ {
		chars[i] = binary.LittleEndian.Uint16(nameData[i*2 : i*2+2])
		if chars[i] == 0 {
			chars = chars[:i]
			break
		}
	}

	return string(utf16.Decode(chars))
}
