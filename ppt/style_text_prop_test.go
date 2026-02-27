package ppt

import (
	"encoding/binary"
	"fmt"
	"math/rand"
	"testing"
	"testing/quick"
)

// buildCharProps builds a TextCFRun entry for testing.
// Per [MS-PPT] TextCFException field order:
// masks, fontStyle, fontRef, oldEAFontRef, ansiFontRef, symbolFontRef, fontSize, color, position
func buildCharProps(count int, bold, italic, underline bool, fontIdx int, fontSize uint16, color string) []byte {
	var buf []byte

	// Character count (4 bytes)
	countBytes := make([]byte, 4)
	binary.LittleEndian.PutUint32(countBytes, uint32(count))
	buf = append(buf, countBytes...)

	// Build CFMasks
	var mask uint32
	var flags uint16
	if bold {
		mask |= uint32(cfBold)
		flags |= uint16(cfBold)
	}
	if italic {
		mask |= uint32(cfItalic)
		flags |= uint16(cfItalic)
	}
	if underline {
		mask |= uint32(cfUnderline)
		flags |= uint16(cfUnderline)
	}
	if fontIdx >= 0 {
		mask |= uint32(cfTypeface)
	}
	if fontSize > 0 {
		mask |= uint32(cfSize)
	}
	if color != "" {
		mask |= uint32(cfColor)
	}

	maskBytes := make([]byte, 4)
	binary.LittleEndian.PutUint32(maskBytes, mask)
	buf = append(buf, maskBytes...)

	// fontStyle (2 bytes) - if any style bits set
	if mask&uint32(cfStyleBits) != 0 {
		flagBytes := make([]byte, 2)
		binary.LittleEndian.PutUint16(flagBytes, flags)
		buf = append(buf, flagBytes...)
	}

	// fontRef (2 bytes)
	if fontIdx >= 0 {
		fontBytes := make([]byte, 2)
		binary.LittleEndian.PutUint16(fontBytes, uint16(fontIdx))
		buf = append(buf, fontBytes...)
	}

	// fontSize (2 bytes) - in points, parser multiplies by 100
	if fontSize > 0 {
		sizeBytes := make([]byte, 2)
		binary.LittleEndian.PutUint16(sizeBytes, fontSize/100) // store as points
		buf = append(buf, sizeBytes...)
	}

	// color (4 bytes)
	if color != "" {
		var r, g, b uint8
		fmt.Sscanf(color, "%02X%02X%02X", &r, &g, &b)
		colorVal := uint32(r) | uint32(g)<<8 | uint32(b)<<16
		colorBytes := make([]byte, 4)
		binary.LittleEndian.PutUint32(colorBytes, colorVal)
		buf = append(buf, colorBytes...)
	}

	return buf
}

// buildParaProps builds a TextPFRun entry for testing.
// Per [MS-PPT] TextPFException field order:
// masks, bulletFlags, bulletChar, bulletFontRef, bulletSize, bulletColor,
// textAlignment, lineSpacing, spaceBefore, spaceAfter, leftMargin, indent, ...
func buildParaProps(count int, indentLevel uint8, alignment int, lineSpacing, spaceBefore, spaceAfter int16, hasBullet bool, bulletChar rune) []byte {
	var buf []byte

	// Character count (4 bytes)
	countBytes := make([]byte, 4)
	binary.LittleEndian.PutUint32(countBytes, uint32(count))
	buf = append(buf, countBytes...)

	// Indent level (2 bytes)
	indentBytes := make([]byte, 2)
	binary.LittleEndian.PutUint16(indentBytes, uint16(indentLevel))
	buf = append(buf, indentBytes...)

	// Build PFMasks
	var mask uint32
	if hasBullet {
		mask |= uint32(pfHasBullet)
	}
	if bulletChar != 0 {
		mask |= uint32(pfBulletChar)
	}
	if alignment >= 0 {
		mask |= uint32(pfAlign)
	}
	if lineSpacing != 0 {
		mask |= uint32(pfLineSpacing)
	}
	if spaceBefore != 0 {
		mask |= uint32(pfSpaceBefore)
	}
	if spaceAfter != 0 {
		mask |= uint32(pfSpaceAfter)
	}

	maskBytes := make([]byte, 4)
	binary.LittleEndian.PutUint32(maskBytes, mask)
	buf = append(buf, maskBytes...)

	// bulletFlags (2 bytes) - if hasBullet bit set
	if hasBullet {
		flagBytes := make([]byte, 2)
		binary.LittleEndian.PutUint16(flagBytes, 0x000F)
		buf = append(buf, flagBytes...)
	}

	// bulletChar (2 bytes)
	if bulletChar != 0 {
		charBytes := make([]byte, 2)
		binary.LittleEndian.PutUint16(charBytes, uint16(bulletChar))
		buf = append(buf, charBytes...)
	}

	// textAlignment (2 bytes)
	if alignment >= 0 {
		alignBytes := make([]byte, 2)
		binary.LittleEndian.PutUint16(alignBytes, uint16(alignment))
		buf = append(buf, alignBytes...)
	}

	// lineSpacing (2 bytes)
	if lineSpacing != 0 {
		lsBytes := make([]byte, 2)
		binary.LittleEndian.PutUint16(lsBytes, uint16(lineSpacing))
		buf = append(buf, lsBytes...)
	}

	// spaceBefore (2 bytes)
	if spaceBefore != 0 {
		sbBytes := make([]byte, 2)
		binary.LittleEndian.PutUint16(sbBytes, uint16(spaceBefore))
		buf = append(buf, sbBytes...)
	}

	// spaceAfter (2 bytes)
	if spaceAfter != 0 {
		saBytes := make([]byte, 2)
		binary.LittleEndian.PutUint16(saBytes, uint16(spaceAfter))
		buf = append(buf, saBytes...)
	}

	return buf
}

func TestParseCharacterProps_BoldItalicUnderline(t *testing.T) {
	fonts := []string{"Arial"}
	data := buildCharProps(10, true, true, true, 0, 2400, "FF0000")
	run, consumed, count := parseCharacterProps(data, 0, fonts)
	if count != 10 {
		t.Errorf("expected count 10, got %d", count)
	}
	if consumed == 0 {
		t.Fatal("consumed should be > 0")
	}
	if !run.Bold {
		t.Error("expected bold")
	}
	if !run.Italic {
		t.Error("expected italic")
	}
	if !run.Underline {
		t.Error("expected underline")
	}
	if run.FontName != "Arial" {
		t.Errorf("expected 'Arial', got %q", run.FontName)
	}
	if run.FontSize != 2400 {
		t.Errorf("expected fontSize 2400, got %d", run.FontSize)
	}
	if run.Color != "FF0000" {
		t.Errorf("expected color 'FF0000', got %q", run.Color)
	}
}

func TestParseCharacterProps_NoFlags(t *testing.T) {
	data := buildCharProps(5, false, false, false, -1, 0, "")
	run, _, count := parseCharacterProps(data, 0, nil)
	if count != 5 {
		t.Errorf("expected count 5, got %d", count)
	}
	if run.Bold || run.Italic || run.Underline {
		t.Error("expected no style flags")
	}
	if run.FontName != "" {
		t.Errorf("expected empty font name, got %q", run.FontName)
	}
}

func TestParseCharacterProps_FontOutOfRange(t *testing.T) {
	fonts := []string{"Arial"}
	data := buildCharProps(5, false, false, false, 99, 0, "")
	run, _, _ := parseCharacterProps(data, 0, fonts)
	if run.FontName != "" {
		t.Errorf("expected empty font name for out-of-range index, got %q", run.FontName)
	}
}

func TestParseParagraphProps_Alignment(t *testing.T) {
	data := buildParaProps(10, 0, 1, 0, 0, 0, false, 0) // center aligned
	para, _, count := parseParagraphProps(data, 0, nil)
	if count != 10 {
		t.Errorf("expected count 10, got %d", count)
	}
	if para.Alignment != 1 {
		t.Errorf("expected alignment 1 (center), got %d", para.Alignment)
	}
}

func TestParseParagraphProps_Bullet(t *testing.T) {
	data := buildParaProps(10, 1, -1, 0, 0, 0, true, '•')
	para, _, _ := parseParagraphProps(data, 0, nil)
	if !para.HasBullet {
		t.Error("expected bullet")
	}
	if para.BulletChar != "•" {
		t.Errorf("expected bullet char '•', got %q", para.BulletChar)
	}
	if para.IndentLevel != 1 {
		t.Errorf("expected indent level 1, got %d", para.IndentLevel)
	}
}

func TestParseParagraphProps_Spacing(t *testing.T) {
	data := buildParaProps(10, 0, -1, 120, 50, 30, false, 0)
	para, _, _ := parseParagraphProps(data, 0, nil)
	if para.LineSpacing != 120 {
		t.Errorf("expected lineSpacing 120, got %d", para.LineSpacing)
	}
	if para.SpaceBefore != 50 {
		t.Errorf("expected spaceBefore 50, got %d", para.SpaceBefore)
	}
	if para.SpaceAfter != 30 {
		t.Errorf("expected spaceAfter 30, got %d", para.SpaceAfter)
	}
}

func TestParseStyleTextPropAtom_Empty(t *testing.T) {
	paras, runs := parseStyleTextPropAtom(nil, 0, nil)
	if paras != nil || runs != nil {
		t.Error("expected nil for empty input")
	}
}

// Property test: character props round-trip
func TestProperty_CharacterPropsParsing(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		bold := rng.Intn(2) == 1
		italic := rng.Intn(2) == 1
		underline := rng.Intn(2) == 1
		fontIdx := rng.Intn(5)
		// fontSize in hundredths of a point (must be multiple of 100 for round-trip)
		fontSizePts := uint16(1 + rng.Intn(100))
		fontSize := fontSizePts * 100
		r := uint8(rng.Intn(256))
		g := uint8(rng.Intn(256))
		b := uint8(rng.Intn(256))
		color := fmt.Sprintf("%02X%02X%02X", r, g, b)
		count := 1 + rng.Intn(100)

		fonts := []string{"Font0", "Font1", "Font2", "Font3", "Font4"}

		data := buildCharProps(count, bold, italic, underline, fontIdx, fontSize, color)
		run, consumed, gotCount := parseCharacterProps(data, 0, fonts)

		if gotCount != count {
			t.Logf("count mismatch: expected %d, got %d", count, gotCount)
			return false
		}
		if consumed == 0 {
			t.Log("consumed should be > 0")
			return false
		}
		if run.Bold != bold {
			t.Logf("bold mismatch: expected %v, got %v", bold, run.Bold)
			return false
		}
		if run.Italic != italic {
			t.Logf("italic mismatch: expected %v, got %v", italic, run.Italic)
			return false
		}
		if run.Underline != underline {
			t.Logf("underline mismatch: expected %v, got %v", underline, run.Underline)
			return false
		}
		if run.FontName != fonts[fontIdx] {
			t.Logf("fontName mismatch: expected %q, got %q", fonts[fontIdx], run.FontName)
			return false
		}
		if run.FontSize != fontSize {
			t.Logf("fontSize mismatch: expected %d, got %d", fontSize, run.FontSize)
			return false
		}
		if run.Color != color {
			t.Logf("color mismatch: expected %q, got %q", color, run.Color)
			return false
		}
		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: character props parsing: %v", err)
	}
}

// Property test: paragraph props round-trip
func TestProperty_ParagraphPropsParsing(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		alignment := rng.Intn(4)
		indentLevel := uint8(rng.Intn(5))
		lineSpacing := int16(rng.Intn(500))
		spaceBefore := int16(rng.Intn(500))
		spaceAfter := int16(rng.Intn(500))
		hasBullet := rng.Intn(2) == 1
		bulletChars := []rune{'•', '-', '*', '>', '→'}
		var bulletChar rune
		if hasBullet {
			bulletChar = bulletChars[rng.Intn(len(bulletChars))]
		}
		count := 1 + rng.Intn(100)

		data := buildParaProps(count, indentLevel, alignment, lineSpacing, spaceBefore, spaceAfter, hasBullet, bulletChar)
		para, consumed, gotCount := parseParagraphProps(data, 0, nil)

		if gotCount != count {
			t.Logf("count mismatch: expected %d, got %d", count, gotCount)
			return false
		}
		if consumed == 0 {
			t.Log("consumed should be > 0")
			return false
		}
		if para.Alignment != uint8(alignment) {
			t.Logf("alignment mismatch: expected %d, got %d", alignment, para.Alignment)
			return false
		}
		if para.IndentLevel != indentLevel {
			t.Logf("indentLevel mismatch: expected %d, got %d", indentLevel, para.IndentLevel)
			return false
		}
		if para.LineSpacing != int32(lineSpacing) {
			t.Logf("lineSpacing mismatch: expected %d, got %d", lineSpacing, para.LineSpacing)
			return false
		}
		if para.SpaceBefore != int32(spaceBefore) {
			t.Logf("spaceBefore mismatch: expected %d, got %d", spaceBefore, para.SpaceBefore)
			return false
		}
		if para.SpaceAfter != int32(spaceAfter) {
			t.Logf("spaceAfter mismatch: expected %d, got %d", spaceAfter, para.SpaceAfter)
			return false
		}
		if para.HasBullet != hasBullet {
			t.Logf("hasBullet mismatch: expected %v, got %v", hasBullet, para.HasBullet)
			return false
		}
		if hasBullet && bulletChar != 0 && para.BulletChar != string(bulletChar) {
			t.Logf("bulletChar mismatch: expected %q, got %q", string(bulletChar), para.BulletChar)
			return false
		}
		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: paragraph props parsing: %v", err)
	}
}
