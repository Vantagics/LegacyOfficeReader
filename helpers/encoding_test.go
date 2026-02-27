package helpers

import (
	"encoding/binary"
	"testing/quick"
	"unicode"
	"unicode/utf16"
	"testing"
)

// Feature: doc-ppt-reader, Property 1: UTF-16LE encoding round-trip
// **Validates: Requirements 7.1, 7.3, 2.6, 4.5**
func TestUTF16LERoundTrip(t *testing.T) {
	config := &quick.Config{MaxCount: 100}
	f := func(s string) bool {
		// Filter out surrogate code points which are invalid in UTF-16 encoding
		filtered := filterSurrogates(s)
		encoded := utf16.Encode([]rune(filtered))
		buf := make([]byte, len(encoded)*2)
		for i, v := range encoded {
			binary.LittleEndian.PutUint16(buf[2*i:], v)
		}
		decoded := DecodeUTF16LE(buf)
		return decoded == filtered
	}
	if err := quick.Check(f, config); err != nil {
		t.Error(err)
	}
}

// filterSurrogates removes lone surrogate code points (U+D800–U+DFFF) from a string,
// as they are not valid Unicode scalar values and cannot round-trip through UTF-16 encoding.
func filterSurrogates(s string) string {
	runes := make([]rune, 0, len(s))
	for _, r := range s {
		if !unicode.Is(unicode.Cs, r) {
			runes = append(runes, r)
		}
	}
	return string(runes)
}

// Feature: doc-ppt-reader, Property 2: ANSI encoding round-trip
// **Validates: Requirements 7.2, 7.4, 2.5, 4.6**
func TestANSIRoundTrip(t *testing.T) {
	config := &quick.Config{MaxCount: 100}
	f := func(data []byte) bool {
		decoded := DecodeANSI(data)
		runes := []rune(decoded)
		if len(runes) != len(data) {
			return false
		}
		for i, r := range runes {
			if byte(r) != data[i] {
				return false
			}
		}
		return true
	}
	if err := quick.Check(f, config); err != nil {
		t.Error(err)
	}
}
