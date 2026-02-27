package ppt

import (
	"encoding/binary"
	"fmt"
)

// StyleTextPropAtom record type
const rtStyleTextPropAtom = 0x0FA1 // 4001

// PFMasks - Paragraph property mask bits per [MS-PPT] 2.9.20
const (
	pfHasBullet      = 0x0001 // bit 0
	pfBulletHasFont  = 0x0002 // bit 1
	pfBulletHasColor = 0x0004 // bit 2
	pfBulletHasSize  = 0x0008 // bit 3
	pfBulletFont     = 0x0010 // bit 4
	pfBulletColor    = 0x0020 // bit 5
	pfBulletSize     = 0x0040 // bit 6
	pfBulletChar     = 0x0080 // bit 7
	pfLeftMargin     = 0x0100 // bit 8
	// bit 9 unused
	pfIndent       = 0x0400 // bit 10
	pfAlign        = 0x0800 // bit 11
	pfLineSpacing  = 0x1000 // bit 12
	pfSpaceBefore  = 0x2000 // bit 13
	pfSpaceAfter   = 0x4000 // bit 14
	pfDefaultTab   = 0x8000 // bit 15
	pfFontAlign    = 0x10000 // bit 16
	pfCharWrap     = 0x20000 // bit 17
	pfWordWrap     = 0x40000 // bit 18
	pfOverflow     = 0x80000 // bit 19
	pfTabStops     = 0x100000 // bit 20
	pfTextDirection = 0x200000 // bit 21
)

// CFMasks - Character property mask bits per [MS-PPT] 2.9.2
const (
	cfBold      = 0x0001 // bit 0
	cfItalic    = 0x0002 // bit 1
	cfUnderline = 0x0004 // bit 2
	// bit 3 unused
	cfShadow  = 0x0010 // bit 4
	cfFehint  = 0x0020 // bit 5
	// bit 6 unused
	cfKumi    = 0x0080 // bit 7
	// bit 8 unused
	cfEmboss  = 0x0200 // bit 9
	cfHasStyle = 0x3C00 // bits 10-13 (fHasStyle, 4 bits)
	// bits 14-15 unused
	cfTypeface       = 0x10000  // bit 16 - fontRef exists
	cfSize           = 0x20000  // bit 17 - fontSize exists
	cfColor          = 0x40000  // bit 18 - color exists
	cfPosition       = 0x80000  // bit 19 - position exists
	// bit 20 pp10ext
	cfOldEATypeface  = 0x200000 // bit 21
	cfAnsiTypeface   = 0x400000 // bit 22
	cfSymbolTypeface = 0x800000 // bit 23
)

// cfStyleBits is the union of all bits that trigger fontStyle field presence
const cfStyleBits = cfBold | cfItalic | cfUnderline | cfShadow | cfFehint | cfKumi | cfEmboss | cfHasStyle

// paraWithCount pairs a SlideParagraph with its character count from the TextPFRun.
type paraWithCount struct {
	Para  SlideParagraph
	Count int
}

// runWithCount pairs a SlideTextRun with its character count from the TextCFRun.
type runWithCount struct {
	Run   SlideTextRun
	Count int
}

// parseStyleTextPropAtom parses paragraph-level and character-level formatting
// from a StyleTextPropAtom data block. textLen is the character count of the
// associated text. fonts is the font index table from FontCollection.
// Returns paragraph properties and character properties as separate slices.
func parseStyleTextPropAtom(data []byte, textLen int, fonts []string) ([]SlideParagraph, []SlideTextRun) {
	if len(data) == 0 || textLen == 0 {
		return nil, nil
	}

	pos := 0
	dataLen := len(data)

	// Parse paragraph properties (TextPFRun array)
	var paras []paraWithCount
	charsCovered := 0
	for charsCovered < textLen && pos+4 <= dataLen {
		para, consumed, count := parseParagraphProps(data, pos, fonts)
		if consumed == 0 {
			break
		}
		paras = append(paras, paraWithCount{Para: para, Count: count})
		pos += consumed
		charsCovered += count
	}

	// Parse character properties (TextCFRun array)
	var runs []runWithCount
	charsCovered = 0
	for charsCovered < textLen && pos+4 <= dataLen {
		run, consumed, count := parseCharacterProps(data, pos, fonts)
		if consumed == 0 {
			break
		}
		runs = append(runs, runWithCount{Run: run, Count: count})
		pos += consumed
		charsCovered += count
	}

	// Convert to simple slices for backward compatibility
	simplePara := make([]SlideParagraph, len(paras))
	for i, p := range paras {
		simplePara[i] = p.Para
	}
	simpleRuns := make([]SlideTextRun, len(runs))
	for i, r := range runs {
		simpleRuns[i] = r.Run
	}

	return simplePara, simpleRuns
}

// parseStyleTextPropAtomWithCounts is like parseStyleTextPropAtom but returns
// character counts alongside each paragraph and run entry.
func parseStyleTextPropAtomWithCounts(data []byte, textLen int, fonts []string) ([]paraWithCount, []runWithCount) {
	if len(data) == 0 || textLen == 0 {
		return nil, nil
	}

	pos := 0
	dataLen := len(data)

	var paras []paraWithCount
	charsCovered := 0
	for charsCovered < textLen && pos+4 <= dataLen {
		para, consumed, count := parseParagraphProps(data, pos, fonts)
		if consumed == 0 {
			break
		}
		paras = append(paras, paraWithCount{Para: para, Count: count})
		pos += consumed
		charsCovered += count
	}

	var runs []runWithCount
	charsCovered = 0
	for charsCovered < textLen && pos+4 <= dataLen {
		run, consumed, count := parseCharacterProps(data, pos, fonts)
		if consumed == 0 {
			break
		}
		runs = append(runs, runWithCount{Run: run, Count: count})
		pos += consumed
		charsCovered += count
	}

	return paras, runs
}

// parseParagraphProps parses a single TextPFRun entry from data at pos.
// Per [MS-PPT] TextPFRun: count (4 bytes) + indentLevel (2 bytes) + TextPFException.
// TextPFException: masks (4 bytes) + optional fields based on mask bits.
func parseParagraphProps(data []byte, pos int, fonts []string) (SlideParagraph, int, int) {
	para := SlideParagraph{}
	dataLen := len(data)
	start := pos

	if pos+4 > dataLen {
		return para, 0, 0
	}

	// Read character count for this paragraph
	count := int(binary.LittleEndian.Uint32(data[pos : pos+4]))
	pos += 4

	// Read indent level (uint16)
	if pos+2 > dataLen {
		return para, pos - start, count
	}
	indentLevel := binary.LittleEndian.Uint16(data[pos : pos+2])
	if indentLevel <= 4 {
		para.IndentLevel = uint8(indentLevel)
	}
	pos += 2

	// Read paragraph property mask (uint32) - PFMasks
	if pos+4 > dataLen {
		return para, pos - start, count
	}
	mask := binary.LittleEndian.Uint32(data[pos : pos+4])
	pos += 4

	// bulletFlags (2 bytes) - exists if any of hasBullet, bulletHasFont, bulletHasColor, bulletHasSize
	if mask&uint32(pfHasBullet|pfBulletHasFont|pfBulletHasColor|pfBulletHasSize) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		bulletFlags := binary.LittleEndian.Uint16(data[pos : pos+2])
		para.HasBullet = bulletFlags&0x0001 != 0
		pos += 2
	}

	// bulletChar (2 bytes)
	if mask&uint32(pfBulletChar) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		ch := binary.LittleEndian.Uint16(data[pos : pos+2])
		para.BulletChar = string(rune(ch))
		pos += 2
	}

	// bulletFontRef (2 bytes)
	if mask&uint32(pfBulletFont) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		fontIdx := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		if fontIdx >= 0 && fontIdx < len(fonts) {
			para.BulletFont = fonts[fontIdx]
		}
		pos += 2
	}

	// bulletSize (2 bytes)
	if mask&uint32(pfBulletSize) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		para.BulletSize = int16(binary.LittleEndian.Uint16(data[pos : pos+2]))
		pos += 2
	}

	// bulletColor (4 bytes)
	if mask&uint32(pfBulletColor) != 0 {
		if pos+4 > dataLen {
			return para, pos - start, count
		}
		colorVal := binary.LittleEndian.Uint32(data[pos : pos+4])
		r := uint8(colorVal & 0xFF)
		g := uint8((colorVal >> 8) & 0xFF)
		b := uint8((colorVal >> 16) & 0xFF)
		para.BulletColor = fmt.Sprintf("%02X%02X%02X", r, g, b)
		pos += 4
	}

	// textAlignment (2 bytes)
	if mask&uint32(pfAlign) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		align := binary.LittleEndian.Uint16(data[pos : pos+2])
		if align <= 3 {
			para.Alignment = uint8(align)
		}
		pos += 2
	}

	// lineSpacing (2 bytes)
	if mask&uint32(pfLineSpacing) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		para.LineSpacing = int32(int16(binary.LittleEndian.Uint16(data[pos : pos+2])))
		pos += 2
	}

	// spaceBefore (2 bytes)
	if mask&uint32(pfSpaceBefore) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		para.SpaceBefore = int32(int16(binary.LittleEndian.Uint16(data[pos : pos+2])))
		pos += 2
	}

	// spaceAfter (2 bytes)
	if mask&uint32(pfSpaceAfter) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		para.SpaceAfter = int32(int16(binary.LittleEndian.Uint16(data[pos : pos+2])))
		pos += 2
	}

	// leftMargin (2 bytes)
	if mask&uint32(pfLeftMargin) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		para.LeftMargin = int32(int16(binary.LittleEndian.Uint16(data[pos : pos+2])))
		pos += 2
	}

	// indent (2 bytes)
	if mask&uint32(pfIndent) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		para.Indent = int32(int16(binary.LittleEndian.Uint16(data[pos : pos+2])))
		pos += 2
	}

	// defaultTabSize (2 bytes)
	if mask&uint32(pfDefaultTab) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		pos += 2
	}

	// tabStops (variable) - starts with uint16 count, then count * 4 bytes
	if mask&uint32(pfTabStops) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		tabCount := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		pos += 2
		skip := tabCount * 4
		if pos+skip > dataLen {
			return para, pos - start, count
		}
		pos += skip
	}

	// fontAlign (2 bytes)
	if mask&uint32(pfFontAlign) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		pos += 2
	}

	// wrapFlags (2 bytes) - exists if any of charWrap, wordWrap, overflow
	if mask&uint32(pfCharWrap|pfWordWrap|pfOverflow) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		pos += 2
	}

	// textDirection (2 bytes)
	if mask&uint32(pfTextDirection) != 0 {
		if pos+2 > dataLen {
			return para, pos - start, count
		}
		pos += 2
	}

	return para, pos - start, count
}

// parseCharacterProps parses a single TextCFRun entry from data at pos.
// Per [MS-PPT] TextCFRun: count (4 bytes) + TextCFException.
// TextCFException: masks (4 bytes) + optional fields per CFMasks.
func parseCharacterProps(data []byte, pos int, fonts []string) (SlideTextRun, int, int) {
	run := SlideTextRun{}
	dataLen := len(data)
	start := pos

	if pos+4 > dataLen {
		return run, 0, 0
	}

	// Read character count
	count := int(binary.LittleEndian.Uint32(data[pos : pos+4]))
	pos += 4

	// Read character property mask (uint32) - CFMasks
	if pos+4 > dataLen {
		return run, pos - start, count
	}
	mask := binary.LittleEndian.Uint32(data[pos : pos+4])
	pos += 4

	// fontStyle (2 bytes) - exists if any style-related bits are set
	if mask&uint32(cfStyleBits) != 0 {
		if pos+2 > dataLen {
			return run, pos - start, count
		}
		flags := binary.LittleEndian.Uint16(data[pos : pos+2])
		run.Bold = flags&uint16(cfBold) != 0
		run.Italic = flags&uint16(cfItalic) != 0
		run.Underline = flags&uint16(cfUnderline) != 0
		pos += 2
	}

	// fontRef (2 bytes) - exists if masks.typeface (bit 16) is set
	if mask&uint32(cfTypeface) != 0 {
		if pos+2 > dataLen {
			return run, pos - start, count
		}
		fontIdx := int(binary.LittleEndian.Uint16(data[pos : pos+2]))
		if fontIdx >= 0 && fontIdx < len(fonts) {
			run.FontName = fonts[fontIdx]
		}
		pos += 2
	}

	// oldEAFontRef (2 bytes) - exists if masks.oldEATypeface (bit 21) is set
	if mask&uint32(cfOldEATypeface) != 0 {
		if pos+2 > dataLen {
			return run, pos - start, count
		}
		pos += 2
	}

	// ansiFontRef (2 bytes) - exists if masks.ansiTypeface (bit 22) is set
	if mask&uint32(cfAnsiTypeface) != 0 {
		if pos+2 > dataLen {
			return run, pos - start, count
		}
		pos += 2
	}

	// symbolFontRef (2 bytes) - exists if masks.symbolTypeface (bit 23) is set
	if mask&uint32(cfSymbolTypeface) != 0 {
		if pos+2 > dataLen {
			return run, pos - start, count
		}
		pos += 2
	}

	// fontSize (2 bytes) - exists if masks.size (bit 17) is set
	if mask&uint32(cfSize) != 0 {
		if pos+2 > dataLen {
			return run, pos - start, count
		}
		// fontSize is in points (not centipoints) per [MS-PPT]
		// OOXML expects hundredths of a point, so multiply by 100
		run.FontSize = binary.LittleEndian.Uint16(data[pos : pos+2]) * 100
		pos += 2
	}

	// color (4 bytes) - exists if masks.color (bit 18) is set
	if mask&uint32(cfColor) != 0 {
		if pos+4 > dataLen {
			return run, pos - start, count
		}
		colorVal := binary.LittleEndian.Uint32(data[pos : pos+4])
		run.ColorRaw = colorVal
		r := uint8(colorVal & 0xFF)
		g := uint8((colorVal >> 8) & 0xFF)
		b := uint8((colorVal >> 16) & 0xFF)
		run.Color = fmt.Sprintf("%02X%02X%02X", r, g, b)
		pos += 4
	}

	// position (2 bytes) - exists if masks.position (bit 19) is set
	if mask&uint32(cfPosition) != 0 {
		if pos+2 > dataLen {
			return run, pos - start, count
		}
		pos += 2
	}

	return run, pos - start, count
}
