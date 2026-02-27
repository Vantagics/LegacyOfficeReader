package doc

import (
	"encoding/binary"
	"unicode/utf16"
)

// parseSttbfFfn parses the font name table (SttbfFfn) from the Table stream.
// Returns a slice of font names indexed by font index.
//
// Per [MS-DOC] 2.9.263, SttbfFfn is an STTB where each "string" is an FFN structure.
// The FFN structure (per [MS-DOC] 2.9.73) contains:
//   Byte 0: cbFfnM1 - size of remaining FFN data minus 1
//   Byte 1: flags (prq, fTrueType, ff)
//   Bytes 2-3: wWeight
//   Byte 4: chs (charset)
//   Byte 5: ixchSzAlt
//   Bytes 6-15: PANOSE (10 bytes)
//   Bytes 16-39: FONTSIGNATURE (24 bytes) - present in Word 97+
//   Byte 40+: xszFfn - font name as null-terminated UTF-16LE string
func parseSttbfFfn(tableData []byte, fc, lcb uint32) []string {
	if lcb == 0 {
		return nil
	}
	if uint64(fc)+uint64(lcb) > uint64(len(tableData)) {
		return nil
	}

	data := tableData[fc : fc+lcb]
	if len(data) < 4 {
		return nil
	}

	// SttbfFfn header: fExtend(2) + cData(2) + cbExtra(2)
	// fExtend should be 0xFFFF for extended STTB
	pos := 0
	fExtend := binary.LittleEndian.Uint16(data[0:2])
	var cData uint16
	if fExtend == 0xFFFF {
		if len(data) < 6 {
			return nil
		}
		cData = binary.LittleEndian.Uint16(data[2:4])
		// cbExtra at data[4:6]
		pos = 6
	} else {
		cData = fExtend
		pos = 4
	}

	fonts := make([]string, 0, int(cData))

	for i := 0; i < int(cData); i++ {
		if pos >= len(data) {
			break
		}

		// Each FFN entry: cbFfnM1 (1 byte) = total size of FFN data - 1
		cbFfnM1 := int(data[pos])
		entryStart := pos + 1
		entryEnd := entryStart + cbFfnM1
		if entryEnd > len(data) {
			break
		}

		entry := data[entryStart:entryEnd]
		pos = entryEnd

		// Try to extract font name from the FFN entry.
		// The font name (xszFfn) is a null-terminated UTF-16LE string.
		// Its position depends on whether FONTSIGNATURE is present.
		//
		// FFN fixed part:
		//   Byte 0: flags
		//   Bytes 1-2: wWeight
		//   Byte 3: chs
		//   Byte 4: ixchSzAlt
		//   Bytes 5-14: PANOSE (10 bytes)
		//   Bytes 15-38: FONTSIGNATURE (24 bytes) - optional
		//   After that: xszFfn (UTF-16LE null-terminated)
		//
		// We try offset 39 first (with FONTSIGNATURE), then 15 (without).
		name := ""
		if len(entry) > 39 {
			name = decodeUTF16NullTerm(entry[39:])
		}
		if !isValidFontName(name) && len(entry) > 15 {
			name = decodeUTF16NullTerm(entry[15:])
		}
		if !isValidFontName(name) {
			// Last resort: scan for the first valid UTF-16LE string
			name = scanForFontName(entry)
		}
		fonts = append(fonts, name)
	}

	return fonts
}

// scanForFontName tries to find a font name by scanning the FFN entry
// for a sequence of valid UTF-16LE characters.
func scanForFontName(entry []byte) string {
	// Try various offsets where the font name might start
	for offset := 39; offset >= 5; offset-- {
		if offset+2 > len(entry) {
			continue
		}
		name := decodeUTF16NullTerm(entry[offset:])
		if isValidFontName(name) {
			return name
		}
	}
	return ""
}

// decodeUTF16NullTerm decodes a null-terminated UTF-16LE string.
func decodeUTF16NullTerm(data []byte) string {
	var u16 []uint16
	for i := 0; i+1 < len(data); i += 2 {
		ch := binary.LittleEndian.Uint16(data[i:])
		if ch == 0 {
			break
		}
		u16 = append(u16, ch)
	}
	return string(utf16.Decode(u16))
}

// isValidFontName checks if a decoded font name looks reasonable.
func isValidFontName(name string) bool {
	if name == "" || len(name) > 100 {
		return false
	}
	runes := []rune(name)
	first := runes[0]
	if first < 0x20 || first == 0xFFFD {
		return false
	}
	for _, r := range runes {
		if r < 0x20 || r == 0xFFFD {
			return false
		}
	}
	return true
}
