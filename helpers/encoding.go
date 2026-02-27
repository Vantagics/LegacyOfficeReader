package helpers

import (
	"encoding/binary"
	"unicode/utf16"

	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/encoding/traditionalchinese"
	"golang.org/x/text/transform"
)

// DecodeUTF16LE converts a UTF-16LE encoded byte sequence to a Go UTF-8 string.
func DecodeUTF16LE(data []byte) string {
	u16s := make([]uint16, len(data)/2)
	for i := range u16s {
		u16s[i] = binary.LittleEndian.Uint16(data[2*i:])
	}
	return string(utf16.Decode(u16s))
}

// DecodeANSI converts an ANSI (Latin-1) encoded byte sequence to a Go UTF-8 string.
// Each byte is mapped directly to its corresponding Unicode code point.
func DecodeANSI(data []byte) string {
	runes := make([]rune, len(data))
	for i, b := range data {
		runes[i] = rune(b)
	}
	return string(runes)
}

// DecodeWithCodepage decodes a byte sequence using the specified Windows codepage.
// Falls back to Latin-1 if the codepage is not supported.
func DecodeWithCodepage(data []byte, codepage uint16) string {
	var decoder *transform.Reader
	switch codepage {
	case 936, 54936: // GBK / GB18030
		decoded, err := simplifiedchinese.GBK.NewDecoder().Bytes(data)
		if err == nil {
			return string(decoded)
		}
	case 950: // Big5
		decoded, err := traditionalchinese.Big5.NewDecoder().Bytes(data)
		if err == nil {
			return string(decoded)
		}
	case 932: // Shift-JIS
		decoded, err := japanese.ShiftJIS.NewDecoder().Bytes(data)
		if err == nil {
			return string(decoded)
		}
	case 949: // EUC-KR
		decoded, err := korean.EUCKR.NewDecoder().Bytes(data)
		if err == nil {
			return string(decoded)
		}
	case 1252, 0: // Windows-1252 (Western European) or default
		decoded, err := charmap.Windows1252.NewDecoder().Bytes(data)
		if err == nil {
			return string(decoded)
		}
	case 1251: // Windows-1251 (Cyrillic)
		decoded, err := charmap.Windows1251.NewDecoder().Bytes(data)
		if err == nil {
			return string(decoded)
		}
	case 1250: // Windows-1250 (Central European)
		decoded, err := charmap.Windows1250.NewDecoder().Bytes(data)
		if err == nil {
			return string(decoded)
		}
	}
	_ = decoder
	// Fallback to Latin-1
	return DecodeANSI(data)
}
