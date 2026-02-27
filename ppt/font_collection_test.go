package ppt

import (
	"encoding/binary"
	"math/rand"
	"testing"
	"testing/quick"
	"unicode/utf16"
)

// buildFontEntityAtom creates a FontEntityAtom record with the given font name.
func buildFontEntityAtom(name string) []byte {
	// FontEntityAtom: recVer=0, recInstance=0, recType=0x0FB7
	// lfFaceName: 64 bytes UTF-16LE, null-terminated
	// Remaining fields: lfCharSet(1) + lfFlags(1) + lfPitchAndFamily(1) + padding = at least 4 bytes
	nameRunes := utf16.Encode([]rune(name))
	var nameBytes [64]byte
	for i, ch := range nameRunes {
		if i >= 32 {
			break
		}
		binary.LittleEndian.PutUint16(nameBytes[i*2:], ch)
	}

	bodyLen := 64 + 4 // lfFaceName + minimal trailing fields
	body := make([]byte, bodyLen)
	copy(body, nameBytes[:])

	// Build record header
	header := make([]byte, 8)
	binary.LittleEndian.PutUint16(header[0:], 0x0000) // recVer=0, recInstance=0
	binary.LittleEndian.PutUint16(header[2:], rtFontEntityAtom)
	binary.LittleEndian.PutUint32(header[4:], uint32(bodyLen))

	return append(header, body...)
}

// buildFontCollection wraps FontEntityAtom records in a FontCollection container.
func buildFontCollection(atoms [][]byte) []byte {
	var body []byte
	for _, atom := range atoms {
		body = append(body, atom...)
	}

	header := make([]byte, 8)
	binary.LittleEndian.PutUint16(header[0:], 0x000F) // recVer=0xF (container)
	binary.LittleEndian.PutUint16(header[2:], rtFontCollection)
	binary.LittleEndian.PutUint32(header[4:], uint32(len(body)))

	return append(header, body...)
}

func TestParseFontCollection_Empty(t *testing.T) {
	// No FontCollection in data
	data := make([]byte, 0)
	fonts := parseFontCollection(data)
	if len(fonts) != 0 {
		t.Errorf("expected empty fonts, got %d", len(fonts))
	}
}

func TestParseFontCollection_EmptyContainer(t *testing.T) {
	// FontCollection container with no FontEntityAtom records
	data := buildFontCollection(nil)
	fonts := parseFontCollection(data)
	if len(fonts) != 0 {
		t.Errorf("expected empty fonts, got %d", len(fonts))
	}
}

func TestParseFontCollection_SingleFont(t *testing.T) {
	atom := buildFontEntityAtom("Arial")
	data := buildFontCollection([][]byte{atom})
	fonts := parseFontCollection(data)
	if len(fonts) != 1 {
		t.Fatalf("expected 1 font, got %d", len(fonts))
	}
	if fonts[0] != "Arial" {
		t.Errorf("expected 'Arial', got %q", fonts[0])
	}
}

func TestParseFontCollection_MultipleFonts(t *testing.T) {
	names := []string{"Arial", "Times New Roman", "Courier New"}
	var atoms [][]byte
	for _, name := range names {
		atoms = append(atoms, buildFontEntityAtom(name))
	}
	data := buildFontCollection(atoms)
	fonts := parseFontCollection(data)
	if len(fonts) != len(names) {
		t.Fatalf("expected %d fonts, got %d", len(names), len(fonts))
	}
	for i, name := range names {
		if fonts[i] != name {
			t.Errorf("font[%d]: expected %q, got %q", i, name, fonts[i])
		}
	}
}

func TestParseFontCollection_ChineseFont(t *testing.T) {
	atom := buildFontEntityAtom("宋体")
	data := buildFontCollection([][]byte{atom})
	fonts := parseFontCollection(data)
	if len(fonts) != 1 {
		t.Fatalf("expected 1 font, got %d", len(fonts))
	}
	if fonts[0] != "宋体" {
		t.Errorf("expected '宋体', got %q", fonts[0])
	}
}

// Feature: ppt-to-pptx-format-conversion, Property 1: FontCollection parsing
// For any valid FontCollection record sequence containing N FontEntityAtom records,
// parseFontCollection should return a slice of length N with matching font names.
func TestProperty_FontCollectionParsing(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(numFonts uint8, seed int64) bool {
		n := int(numFonts) % 10 // 0-9 fonts
		rng := rand.New(rand.NewSource(seed))

		names := make([]string, n)
		var atoms [][]byte
		for i := 0; i < n; i++ {
			// Generate random font name (ASCII, 1-20 chars)
			nameLen := 1 + rng.Intn(20)
			nameBytes := make([]byte, nameLen)
			for j := range nameBytes {
				nameBytes[j] = byte('A' + rng.Intn(26))
			}
			names[i] = string(nameBytes)
			atoms = append(atoms, buildFontEntityAtom(names[i]))
		}

		data := buildFontCollection(atoms)
		fonts := parseFontCollection(data)

		if len(fonts) != n {
			t.Logf("expected %d fonts, got %d", n, len(fonts))
			return false
		}
		for i := 0; i < n; i++ {
			if fonts[i] != names[i] {
				t.Logf("font[%d]: expected %q, got %q", i, names[i], fonts[i])
				return false
			}
		}
		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: FontCollection parsing: %v", err)
	}
}
