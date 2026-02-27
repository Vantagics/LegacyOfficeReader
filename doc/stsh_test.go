package doc

import (
	"encoding/binary"
	"strings"
	"testing"
	"unicode/utf16"
)

// buildSTSHBlob constructs a minimal valid STSH binary blob.
// cbSTDBaseInFile is the fixed base size of each STD entry.
// styles is a list of (name, stk, istdBase) tuples.
// stk: 1=paragraph, 2=character.
func buildSTSHBlob(cbSTDBaseInFile uint16, styles []struct {
	name     string
	stk      uint8
	istdBase uint16
}) []byte {
	cstd := uint16(len(styles))

	// Stshi header: cbStshi (uint16) + cstd (uint16) + cbSTDBaseInFile (uint16) + padding
	// cbStshi = size of Stshi fixed portion (not counting the cbStshi field itself).
	// Minimum Stshi has cstd (2 bytes) + cbSTDBaseInFile (2 bytes) = 4 bytes.
	// We'll use cbStshi = 4 for simplicity.
	cbStshi := uint16(4)

	var buf []byte

	// Write cbStshi
	tmp := make([]byte, 2)
	binary.LittleEndian.PutUint16(tmp, cbStshi)
	buf = append(buf, tmp...)

	// Write cstd
	binary.LittleEndian.PutUint16(tmp, cstd)
	buf = append(buf, tmp...)

	// Write cbSTDBaseInFile
	binary.LittleEndian.PutUint16(tmp, cbSTDBaseInFile)
	buf = append(buf, tmp...)

	// Now write each STD entry
	for _, s := range styles {
		stdData := buildSTDEntry(s.name, s.stk, s.istdBase, cbSTDBaseInFile)
		// Write cbStd (uint16) = length of stdData
		cbStd := uint16(len(stdData))
		binary.LittleEndian.PutUint16(tmp, cbStd)
		buf = append(buf, tmp...)
		buf = append(buf, stdData...)
	}

	return buf
}

// buildSTDEntry constructs a single STD entry's data (without the cbStd prefix).
func buildSTDEntry(name string, stk uint8, istdBase uint16, cbSTDBaseInFile uint16) []byte {
	// Fixed base: per [MS-DOC] 2.9.260 Stdf:
	// word0 (bytes 0-1): bits 0-11 = sti (0 for simplicity), bits 12-15 = flags
	// word1 (bytes 2-3): bits 0-3 = stk, bits 4-15 = istdBase
	base := make([]byte, cbSTDBaseInFile)

	// word0: sti=0, flags=0
	binary.LittleEndian.PutUint16(base[0:2], 0)

	// word1: stk in bits 0-3, istdBase in bits 4-15
	if cbSTDBaseInFile >= 4 {
		word1 := uint16(stk) | ((istdBase & 0x0FFF) << 4)
		binary.LittleEndian.PutUint16(base[2:4], word1)
	}

	// Style name after the fixed base
	u16 := utf16.Encode([]rune(name))
	nameLen := uint16(len(u16))

	nameBuf := make([]byte, 2+len(u16)*2+2) // length + chars + null terminator
	binary.LittleEndian.PutUint16(nameBuf[0:2], nameLen)
	for i, c := range u16 {
		binary.LittleEndian.PutUint16(nameBuf[2+i*2:2+i*2+2], c)
	}
	// null terminator (2 zero bytes) already present from make

	var result []byte
	result = append(result, base...)
	result = append(result, nameBuf...)
	return result
}

func TestParseSTSH_ValidData(t *testing.T) {
	styles := []struct {
		name     string
		stk      uint8
		istdBase uint16
	}{
		{"Normal", 1, 0x0FFF},       // paragraph style, no base
		{"heading 1", 1, 0},         // paragraph style, based on Normal (index 0)
		{"Default Paragraph Font", 2, 0x0FFF}, // character style, no base
	}

	blob := buildSTSHBlob(10, styles)

	// Place blob in a "table stream" at offset 100
	tableData := make([]byte, 100+len(blob))
	copy(tableData[100:], blob)

	result, err := parseSTSH(tableData, 100, uint32(len(blob)), nil)
	if err != nil {
		t.Fatalf("parseSTSH returned unexpected error: %v", err)
	}

	if len(result) != 3 {
		t.Fatalf("expected 3 styles, got %d", len(result))
	}

	// Check style 0: Normal
	if result[0].name != "Normal" {
		t.Errorf("style[0].name = %q, want %q", result[0].name, "Normal")
	}
	if result[0].styleType != styleTypeParagraph {
		t.Errorf("style[0].styleType = %d, want %d", result[0].styleType, styleTypeParagraph)
	}
	if result[0].istdBase != 0x0FFF {
		t.Errorf("style[0].istdBase = 0x%04X, want 0x0FFF", result[0].istdBase)
	}

	// Check style 1: heading 1
	if result[1].name != "heading 1" {
		t.Errorf("style[1].name = %q, want %q", result[1].name, "heading 1")
	}
	if result[1].styleType != styleTypeParagraph {
		t.Errorf("style[1].styleType = %d, want %d", result[1].styleType, styleTypeParagraph)
	}
	if result[1].istdBase != 0 {
		t.Errorf("style[1].istdBase = %d, want 0", result[1].istdBase)
	}

	// Check style 2: Default Paragraph Font
	if result[2].name != "Default Paragraph Font" {
		t.Errorf("style[2].name = %q, want %q", result[2].name, "Default Paragraph Font")
	}
	if result[2].styleType != styleTypeCharacter {
		t.Errorf("style[2].styleType = %d, want %d", result[2].styleType, styleTypeCharacter)
	}
	if result[2].istdBase != 0x0FFF {
		t.Errorf("style[2].istdBase = 0x%04X, want 0x0FFF", result[2].istdBase)
	}

	// charProps and paraProps should be nil (UPX parsing deferred)
	for i, s := range result {
		if s.charProps != nil {
			t.Errorf("style[%d].charProps should be nil", i)
		}
		if s.paraProps != nil {
			t.Errorf("style[%d].paraProps should be nil", i)
		}
	}
}

func TestParseSTSH_OutOfBounds(t *testing.T) {
	tableData := make([]byte, 50)

	// fc + lcb exceeds tableData length
	_, err := parseSTSH(tableData, 40, 20, nil)
	if err == nil {
		t.Fatal("parseSTSH should return error when data is out of bounds")
	}
	if !strings.Contains(err.Error(), "out of bounds") {
		t.Errorf("error message should contain 'out of bounds', got: %v", err)
	}
}

func TestParseSTSH_EmptySTSH(t *testing.T) {
	tableData := make([]byte, 100)

	// lcb = 0 means empty STSH
	result, err := parseSTSH(tableData, 0, 0, nil)
	if err != nil {
		t.Fatalf("parseSTSH returned unexpected error: %v", err)
	}
	if len(result) != 0 {
		t.Errorf("expected empty slice, got %d styles", len(result))
	}
}
