package doc

import (
	"encoding/binary"
	"fmt"
	"math/rand"
	"testing"
	"testing/quick"
)

// buildChpxSprms constructs a sprm byte sequence from a CharacterFormatting value.
// This is the inverse of parseChpxSprms for the known opcodes.
func buildChpxSprms(cf CharacterFormatting) []byte {
	var buf []byte
	tmp2 := make([]byte, 2)
	tmp4 := make([]byte, 4)

	// sprmCRgFtc0 (0x4A4F) - font index, uint16
	if cf.FontName != "" {
		var fontIdx uint16
		fmt.Sscanf(cf.FontName, "font%d", &fontIdx)
		binary.LittleEndian.PutUint16(tmp2, 0x4A4F)
		buf = append(buf, tmp2...)
		binary.LittleEndian.PutUint16(tmp2, fontIdx)
		buf = append(buf, tmp2...)
	}

	// sprmCHps (0x4A43) - font size, uint16
	if cf.FontSize > 0 {
		binary.LittleEndian.PutUint16(tmp2, 0x4A43)
		buf = append(buf, tmp2...)
		binary.LittleEndian.PutUint16(tmp2, cf.FontSize)
		buf = append(buf, tmp2...)
	}

	// sprmCFBold (0x0835) - toggle, 1 byte
	if cf.Bold {
		binary.LittleEndian.PutUint16(tmp2, 0x0835)
		buf = append(buf, tmp2...)
		buf = append(buf, 1)
	}

	// sprmCFItalic (0x0836) - toggle, 1 byte
	if cf.Italic {
		binary.LittleEndian.PutUint16(tmp2, 0x0836)
		buf = append(buf, tmp2...)
		buf = append(buf, 1)
	}

	// sprmCKul (0x2A3E) - underline, 1 byte
	if cf.Underline > 0 {
		binary.LittleEndian.PutUint16(tmp2, 0x2A3E)
		buf = append(buf, tmp2...)
		buf = append(buf, cf.Underline)
	}

	// sprmCCv (0x6870) - direct RGB, 4 bytes (we use this for color)
	if cf.Color != "" {
		var r, g, b uint8
		fmt.Sscanf(cf.Color, "%02X%02X%02X", &r, &g, &b)
		binary.LittleEndian.PutUint16(tmp2, 0x6870)
		buf = append(buf, tmp2...)
		tmp4[0] = r
		tmp4[1] = g
		tmp4[2] = b
		tmp4[3] = 0
		buf = append(buf, tmp4...)
	}

	// sprmCIstd (0x4A30) - style index, uint16
	if cf.IstdChar > 0 {
		binary.LittleEndian.PutUint16(tmp2, 0x4A30)
		buf = append(buf, tmp2...)
		binary.LittleEndian.PutUint16(tmp2, cf.IstdChar)
		buf = append(buf, tmp2...)
	}

	return buf
}

// **Feature: doc-format-preservation, Property 2: 字符 Sprm 解析**
// **Validates: Requirements 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8**
func TestPropertyChpxSprms(t *testing.T) {
	cfg := &quick.Config{
		MaxCount: 100,
		Rand:     rand.New(rand.NewSource(42)),
	}

	f := func(fontIdx uint16, fontSize uint16, bold, italic bool, underline uint8, r, g, b uint8, istdChar uint16) bool {
		// Constrain fontSize to non-zero (we only encode when > 0)
		if fontSize == 0 {
			fontSize = 1
		}
		// Constrain underline to non-zero
		if underline == 0 {
			underline = 1
		}
		// Constrain istdChar to non-zero
		if istdChar == 0 {
			istdChar = 1
		}

		expected := CharacterFormatting{
			FontName:  fmt.Sprintf("font%d", fontIdx),
			FontSize:  fontSize,
			Bold:      bold,
			Italic:    italic,
			Underline: underline,
			Color:     fmt.Sprintf("%02X%02X%02X", r, g, b),
			IstdChar:  istdChar,
		}

		sprmData := buildChpxSprms(expected)
		result := parseChpxSprms(sprmData, nil, nil)

		if result.FontName != expected.FontName {
			t.Logf("FontName: got %q, want %q", result.FontName, expected.FontName)
			return false
		}
		if result.FontSize != expected.FontSize {
			t.Logf("FontSize: got %d, want %d", result.FontSize, expected.FontSize)
			return false
		}
		if result.Bold != expected.Bold {
			t.Logf("Bold: got %v, want %v", result.Bold, expected.Bold)
			return false
		}
		if result.Italic != expected.Italic {
			t.Logf("Italic: got %v, want %v", result.Italic, expected.Italic)
			return false
		}
		if result.Underline != expected.Underline {
			t.Logf("Underline: got %d, want %d", result.Underline, expected.Underline)
			return false
		}
		if result.Color != expected.Color {
			t.Logf("Color: got %q, want %q", result.Color, expected.Color)
			return false
		}
		if result.IstdChar != expected.IstdChar {
			t.Logf("IstdChar: got %d, want %d", result.IstdChar, expected.IstdChar)
			return false
		}
		return true
	}

	if err := quick.Check(f, cfg); err != nil {
		t.Errorf("Property test failed: %v", err)
	}
}


func TestParseChpxSprms_AllFields(t *testing.T) {
	var buf []byte
	tmp2 := make([]byte, 2)
	tmp4 := make([]byte, 4)

	// sprmCRgFtc0 (0x4A4F) - font index 42
	binary.LittleEndian.PutUint16(tmp2, 0x4A4F)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 42)
	buf = append(buf, tmp2...)

	// sprmCHps (0x4A43) - font size 24 half-points
	binary.LittleEndian.PutUint16(tmp2, 0x4A43)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 24)
	buf = append(buf, tmp2...)

	// sprmCFBold (0x0835) - bold on
	binary.LittleEndian.PutUint16(tmp2, 0x0835)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// sprmCFItalic (0x0836) - italic on
	binary.LittleEndian.PutUint16(tmp2, 0x0836)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// sprmCKul (0x2A3E) - underline type 1 (single)
	binary.LittleEndian.PutUint16(tmp2, 0x2A3E)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// sprmCCv (0x6870) - direct RGB: R=0xAB, G=0xCD, B=0xEF
	binary.LittleEndian.PutUint16(tmp2, 0x6870)
	buf = append(buf, tmp2...)
	tmp4[0] = 0xAB
	tmp4[1] = 0xCD
	tmp4[2] = 0xEF
	tmp4[3] = 0x00
	buf = append(buf, tmp4...)

	// sprmCIstd (0x4A30) - style index 5
	binary.LittleEndian.PutUint16(tmp2, 0x4A30)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 5)
	buf = append(buf, tmp2...)

	result := parseChpxSprms(buf, nil, nil)

	if result.FontName != "font42" {
		t.Errorf("FontName = %q, want %q", result.FontName, "font42")
	}
	if result.FontSize != 24 {
		t.Errorf("FontSize = %d, want 24", result.FontSize)
	}
	if !result.Bold {
		t.Error("Bold = false, want true")
	}
	if !result.Italic {
		t.Error("Italic = false, want true")
	}
	if result.Underline != 1 {
		t.Errorf("Underline = %d, want 1", result.Underline)
	}
	if result.Color != "ABCDEF" {
		t.Errorf("Color = %q, want %q", result.Color, "ABCDEF")
	}
	if result.IstdChar != 5 {
		t.Errorf("IstdChar = %d, want 5", result.IstdChar)
	}
}

func TestParseChpxSprms_UnknownSprm(t *testing.T) {
	var buf []byte
	tmp2 := make([]byte, 2)

	// sprmCFBold (0x0835) - bold on
	binary.LittleEndian.PutUint16(tmp2, 0x0835)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// Unknown sprm with spra=1 (1-byte operand): opcode 0x2800
	binary.LittleEndian.PutUint16(tmp2, 0x2800)
	buf = append(buf, tmp2...)
	buf = append(buf, 0xFF) // unknown operand

	// sprmCFItalic (0x0836) - italic on
	binary.LittleEndian.PutUint16(tmp2, 0x0836)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// Another unknown sprm with spra=2 (2-byte operand): opcode 0x4801
	binary.LittleEndian.PutUint16(tmp2, 0x4801)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 0xBEEF)
	buf = append(buf, tmp2...)

	// sprmCHps (0x4A43) - font size 20
	binary.LittleEndian.PutUint16(tmp2, 0x4A43)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 20)
	buf = append(buf, tmp2...)

	result := parseChpxSprms(buf, nil, nil)

	if !result.Bold {
		t.Error("Bold = false, want true (should survive unknown sprms)")
	}
	if !result.Italic {
		t.Error("Italic = false, want true (should survive unknown sprms)")
	}
	if result.FontSize != 20 {
		t.Errorf("FontSize = %d, want 20 (should survive unknown sprms)", result.FontSize)
	}
}

func TestParsePlcBteChpx_OutOfBounds(t *testing.T) {
	tableData := make([]byte, 50)

	// fc + lcb exceeds tableData length
	_, err := parsePlcBteChpx(nil, tableData, 40, 20, nil, nil, nil)
	if err == nil {
		t.Fatal("parsePlcBteChpx should return error when data is out of bounds")
	}
	if err.Error() != "PlcBteChpx data out of bounds" {
		t.Errorf("unexpected error message: %v", err)
	}
}
