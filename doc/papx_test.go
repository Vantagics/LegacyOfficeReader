package doc

import (
	"encoding/binary"
	"math/rand"
	"testing"
	"testing/quick"
)

// buildPapxSprms constructs a sprm byte sequence from paragraph format values.
// This is the inverse of parsePapxSprms for the known opcodes.
func buildPapxSprms(pf paraFormatRun) []byte {
	var buf []byte
	tmp2 := make([]byte, 2)
	tmp4 := make([]byte, 4)

	// sprmPJc (0x2461) - alignment, 1 byte
	binary.LittleEndian.PutUint16(tmp2, 0x2461)
	buf = append(buf, tmp2...)
	buf = append(buf, pf.props.Alignment)

	// sprmPDxaLeft (0x845E) - left indent, int16
	binary.LittleEndian.PutUint16(tmp2, 0x845E)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, uint16(int16(pf.props.IndentLeft)))
	buf = append(buf, tmp2...)

	// sprmPDxaRight (0x845D) - right indent, int16
	binary.LittleEndian.PutUint16(tmp2, 0x845D)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, uint16(int16(pf.props.IndentRight)))
	buf = append(buf, tmp2...)

	// sprmPDxaLeft1 (0x8460) - first line indent, int16
	binary.LittleEndian.PutUint16(tmp2, 0x8460)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, uint16(int16(pf.props.IndentFirst)))
	buf = append(buf, tmp2...)

	// sprmPDyaBefore (0xA413) - space before, uint16
	binary.LittleEndian.PutUint16(tmp2, 0xA413)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, pf.props.SpaceBefore)
	buf = append(buf, tmp2...)

	// sprmPDyaAfter (0xA414) - space after, uint16
	binary.LittleEndian.PutUint16(tmp2, 0xA414)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, pf.props.SpaceAfter)
	buf = append(buf, tmp2...)

	// sprmPDyaLine (0x6412) - line spacing, 4 bytes
	// Encode based on LineRule:
	// LineRule=0 (auto): fMult=1, dyaLine=LineSpacing
	// LineRule=1 (atLeast): fMult=0, dyaLine=LineSpacing (positive)
	// LineRule=2 (exact): fMult=0, dyaLine=-LineSpacing (negative)
	binary.LittleEndian.PutUint16(tmp2, 0x6412)
	buf = append(buf, tmp2...)
	{
		var dyaLine int16
		var fMult uint16
		switch pf.props.LineRule {
		case 0: // auto
			dyaLine = int16(pf.props.LineSpacing)
			fMult = 1
		case 1: // atLeast
			dyaLine = int16(pf.props.LineSpacing)
			fMult = 0
		case 2: // exact
			dyaLine = -int16(pf.props.LineSpacing)
			fMult = 0
		default:
			dyaLine = int16(pf.props.LineSpacing)
			fMult = 0
		}
		binary.LittleEndian.PutUint16(tmp4[0:2], uint16(dyaLine))
		binary.LittleEndian.PutUint16(tmp4[2:4], fMult)
	}
	buf = append(buf, tmp4...)

	// sprmPIstd (0x4600) - style index, uint16
	binary.LittleEndian.PutUint16(tmp2, 0x4600)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, pf.istd)
	buf = append(buf, tmp2...)

	// sprmPFInTable (0x2416) - in table, 1 byte
	binary.LittleEndian.PutUint16(tmp2, 0x2416)
	buf = append(buf, tmp2...)
	if pf.inTable {
		buf = append(buf, 1)
	} else {
		buf = append(buf, 0)
	}

	// sprmPFTtp (0x2417) - table row end, 1 byte
	binary.LittleEndian.PutUint16(tmp2, 0x2417)
	buf = append(buf, tmp2...)
	if pf.tableRowEnd {
		buf = append(buf, 1)
	} else {
		buf = append(buf, 0)
	}

	// sprmPIlfo (0x460B) - list override index, uint16
	binary.LittleEndian.PutUint16(tmp2, 0x460B)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, pf.ilfo)
	buf = append(buf, tmp2...)

	// sprmPIlvl (0x260A) - list level, 1 byte
	binary.LittleEndian.PutUint16(tmp2, 0x260A)
	buf = append(buf, tmp2...)
	buf = append(buf, pf.ilvl)

	// sprmPFPageBreakBefore (0x2407) - page break before, 1 byte
	binary.LittleEndian.PutUint16(tmp2, 0x2407)
	buf = append(buf, tmp2...)
	if pf.pageBreakBefore {
		buf = append(buf, 1)
	} else {
		buf = append(buf, 0)
	}

	// sprmPOutLvl (0x2640) - outline level, 1 byte
	binary.LittleEndian.PutUint16(tmp2, 0x2640)
	buf = append(buf, tmp2...)
	buf = append(buf, pf.outLvl)

	return buf
}

// **Feature: doc-format-preservation, Property 3: 段落 Sprm 解析**
// **Validates: Requirements 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 5.2, 6.1, 6.3, 7.1, 7.2, 8.2**
func TestPropertyPapxSprms(t *testing.T) {
	cfg := &quick.Config{
		MaxCount: 100,
		Rand:     rand.New(rand.NewSource(42)),
	}

	f := func(
		alignment uint8,
		indentLeft, indentRight, indentFirst int16,
		spaceBefore, spaceAfter uint16,
		lineSpacing int16, lineRule uint8,
		istd uint16,
		inTable, tableRowEnd bool,
		ilfo uint16, ilvl uint8,
		pageBreakBefore bool,
		outLvl uint8,
	) bool {
		// Constrain alignment to valid range 0-3
		alignment = alignment % 4
		// Constrain outLvl to 0-9
		outLvl = outLvl % 10
		// Constrain lineRule to valid values 0-2
		lineRule = lineRule % 3
		// For round-trip testing, LineSpacing must be positive (exact rule negates it internally)
		if lineSpacing < 0 {
			lineSpacing = -lineSpacing
		}
		if lineSpacing == 0 {
			lineSpacing = 1
		}

		expected := paraFormatRun{
			props: ParagraphFormatting{
				Alignment:   alignment,
				IndentLeft:  int32(indentLeft),
				IndentRight: int32(indentRight),
				IndentFirst: int32(indentFirst),
				SpaceBefore: spaceBefore,
				SpaceAfter:  spaceAfter,
				LineSpacing: int32(lineSpacing),
				LineRule:    lineRule,
			},
			istd:            istd,
			inTable:         inTable,
			tableRowEnd:     tableRowEnd,
			ilfo:            ilfo,
			ilvl:            ilvl,
			pageBreakBefore: pageBreakBefore,
			outLvl:          outLvl,
		}

		sprmData := buildPapxSprms(expected)
		result := parsePapxSprms(sprmData)

		if result.props.Alignment != expected.props.Alignment {
			t.Logf("Alignment: got %d, want %d", result.props.Alignment, expected.props.Alignment)
			return false
		}
		if result.props.IndentLeft != expected.props.IndentLeft {
			t.Logf("IndentLeft: got %d, want %d", result.props.IndentLeft, expected.props.IndentLeft)
			return false
		}
		if result.props.IndentRight != expected.props.IndentRight {
			t.Logf("IndentRight: got %d, want %d", result.props.IndentRight, expected.props.IndentRight)
			return false
		}
		if result.props.IndentFirst != expected.props.IndentFirst {
			t.Logf("IndentFirst: got %d, want %d", result.props.IndentFirst, expected.props.IndentFirst)
			return false
		}
		if result.props.SpaceBefore != expected.props.SpaceBefore {
			t.Logf("SpaceBefore: got %d, want %d", result.props.SpaceBefore, expected.props.SpaceBefore)
			return false
		}
		if result.props.SpaceAfter != expected.props.SpaceAfter {
			t.Logf("SpaceAfter: got %d, want %d", result.props.SpaceAfter, expected.props.SpaceAfter)
			return false
		}
		if result.props.LineSpacing != expected.props.LineSpacing {
			t.Logf("LineSpacing: got %d, want %d", result.props.LineSpacing, expected.props.LineSpacing)
			return false
		}
		if result.props.LineRule != expected.props.LineRule {
			t.Logf("LineRule: got %d, want %d", result.props.LineRule, expected.props.LineRule)
			return false
		}
		if result.istd != expected.istd {
			t.Logf("istd: got %d, want %d", result.istd, expected.istd)
			return false
		}
		if result.inTable != expected.inTable {
			t.Logf("inTable: got %v, want %v", result.inTable, expected.inTable)
			return false
		}
		if result.tableRowEnd != expected.tableRowEnd {
			t.Logf("tableRowEnd: got %v, want %v", result.tableRowEnd, expected.tableRowEnd)
			return false
		}
		if result.ilfo != expected.ilfo {
			t.Logf("ilfo: got %d, want %d", result.ilfo, expected.ilfo)
			return false
		}
		if result.ilvl != expected.ilvl {
			t.Logf("ilvl: got %d, want %d", result.ilvl, expected.ilvl)
			return false
		}
		if result.pageBreakBefore != expected.pageBreakBefore {
			t.Logf("pageBreakBefore: got %v, want %v", result.pageBreakBefore, expected.pageBreakBefore)
			return false
		}
		if result.outLvl != expected.outLvl {
			t.Logf("outLvl: got %d, want %d", result.outLvl, expected.outLvl)
			return false
		}
		return true
	}

	if err := quick.Check(f, cfg); err != nil {
		t.Errorf("Property test failed: %v", err)
	}
}


func TestParsePapxSprms_AllFields(t *testing.T) {
	var buf []byte
	tmp2 := make([]byte, 2)
	tmp4 := make([]byte, 4)

	// sprmPJc (0x2461) - alignment: center (1)
	binary.LittleEndian.PutUint16(tmp2, 0x2461)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// sprmPDxaLeft (0x845E) - left indent: 720 twips
	binary.LittleEndian.PutUint16(tmp2, 0x845E)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, uint16(int16(720)))
	buf = append(buf, tmp2...)

	// sprmPDxaRight (0x845D) - right indent: 360 twips
	binary.LittleEndian.PutUint16(tmp2, 0x845D)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, uint16(int16(360)))
	buf = append(buf, tmp2...)

	// sprmPDxaLeft1 (0x8460) - first line indent: -240 twips (hanging)
	binary.LittleEndian.PutUint16(tmp2, 0x8460)
	buf = append(buf, tmp2...)
	{
		v := int16(-240)
		binary.LittleEndian.PutUint16(tmp2, uint16(v))
	}
	buf = append(buf, tmp2...)

	// sprmPDyaBefore (0xA413) - space before: 120 twips
	binary.LittleEndian.PutUint16(tmp2, 0xA413)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 120)
	buf = append(buf, tmp2...)

	// sprmPDyaAfter (0xA414) - space after: 60 twips
	binary.LittleEndian.PutUint16(tmp2, 0xA414)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 60)
	buf = append(buf, tmp2...)

	// sprmPDyaLine (0x6412) - line spacing: 360, rule: atLeast (fMult=0, dyaLine=360 positive)
	binary.LittleEndian.PutUint16(tmp2, 0x6412)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, uint16(int16(360)))
	tmp4[0] = tmp2[0]
	tmp4[1] = tmp2[1]
	tmp4[2] = 0 // fMult low byte = 0 (not auto)
	tmp4[3] = 0 // fMult high byte = 0
	buf = append(buf, tmp4...)

	// sprmPIstd (0x4600) - style index: 3
	binary.LittleEndian.PutUint16(tmp2, 0x4600)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 3)
	buf = append(buf, tmp2...)

	// sprmPFInTable (0x2416) - in table: true
	binary.LittleEndian.PutUint16(tmp2, 0x2416)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// sprmPFTtp (0x2417) - table row end: true
	binary.LittleEndian.PutUint16(tmp2, 0x2417)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// sprmPIlfo (0x460B) - list override index: 2
	binary.LittleEndian.PutUint16(tmp2, 0x460B)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 2)
	buf = append(buf, tmp2...)

	// sprmPIlvl (0x260A) - list level: 3
	binary.LittleEndian.PutUint16(tmp2, 0x260A)
	buf = append(buf, tmp2...)
	buf = append(buf, 3)

	// sprmPFPageBreakBefore (0x2407) - page break before: true
	binary.LittleEndian.PutUint16(tmp2, 0x2407)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// sprmPOutLvl (0x2640) - outline level: 2
	binary.LittleEndian.PutUint16(tmp2, 0x2640)
	buf = append(buf, tmp2...)
	buf = append(buf, 2)

	result := parsePapxSprms(buf)

	if result.props.Alignment != 1 {
		t.Errorf("Alignment = %d, want 1", result.props.Alignment)
	}
	if result.props.IndentLeft != 720 {
		t.Errorf("IndentLeft = %d, want 720", result.props.IndentLeft)
	}
	if result.props.IndentRight != 360 {
		t.Errorf("IndentRight = %d, want 360", result.props.IndentRight)
	}
	if result.props.IndentFirst != -240 {
		t.Errorf("IndentFirst = %d, want -240", result.props.IndentFirst)
	}
	if result.props.SpaceBefore != 120 {
		t.Errorf("SpaceBefore = %d, want 120", result.props.SpaceBefore)
	}
	if result.props.SpaceAfter != 60 {
		t.Errorf("SpaceAfter = %d, want 60", result.props.SpaceAfter)
	}
	if result.props.LineSpacing != 360 {
		t.Errorf("LineSpacing = %d, want 360", result.props.LineSpacing)
	}
	if result.props.LineRule != 1 {
		t.Errorf("LineRule = %d, want 1", result.props.LineRule)
	}
	if result.istd != 3 {
		t.Errorf("istd = %d, want 3", result.istd)
	}
	if !result.inTable {
		t.Error("inTable = false, want true")
	}
	if !result.tableRowEnd {
		t.Error("tableRowEnd = false, want true")
	}
	if result.ilfo != 2 {
		t.Errorf("ilfo = %d, want 2", result.ilfo)
	}
	if result.ilvl != 3 {
		t.Errorf("ilvl = %d, want 3", result.ilvl)
	}
	if !result.pageBreakBefore {
		t.Error("pageBreakBefore = false, want true")
	}
	if result.outLvl != 2 {
		t.Errorf("outLvl = %d, want 2", result.outLvl)
	}
}

func TestParsePapxSprms_UnknownSprm(t *testing.T) {
	var buf []byte
	tmp2 := make([]byte, 2)

	// sprmPJc (0x2461) - alignment: right (2)
	binary.LittleEndian.PutUint16(tmp2, 0x2461)
	buf = append(buf, tmp2...)
	buf = append(buf, 2)

	// Unknown sprm with spra=1 (1-byte operand): opcode 0x2800
	binary.LittleEndian.PutUint16(tmp2, 0x2800)
	buf = append(buf, tmp2...)
	buf = append(buf, 0xFF) // unknown operand

	// sprmPFInTable (0x2416) - in table: true
	binary.LittleEndian.PutUint16(tmp2, 0x2416)
	buf = append(buf, tmp2...)
	buf = append(buf, 1)

	// Another unknown sprm with spra=2 (2-byte operand): opcode 0x4801
	binary.LittleEndian.PutUint16(tmp2, 0x4801)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 0xBEEF)
	buf = append(buf, tmp2...)

	// sprmPDyaBefore (0xA413) - space before: 200
	binary.LittleEndian.PutUint16(tmp2, 0xA413)
	buf = append(buf, tmp2...)
	binary.LittleEndian.PutUint16(tmp2, 200)
	buf = append(buf, tmp2...)

	result := parsePapxSprms(buf)

	if result.props.Alignment != 2 {
		t.Errorf("Alignment = %d, want 2 (should survive unknown sprms)", result.props.Alignment)
	}
	if !result.inTable {
		t.Error("inTable = false, want true (should survive unknown sprms)")
	}
	if result.props.SpaceBefore != 200 {
		t.Errorf("SpaceBefore = %d, want 200 (should survive unknown sprms)", result.props.SpaceBefore)
	}
}

func TestParsePlcBtePapx_OutOfBounds(t *testing.T) {
	tableData := make([]byte, 50)

	// fc + lcb exceeds tableData length
	_, err := parsePlcBtePapx(nil, tableData, 40, 20, nil)
	if err == nil {
		t.Fatal("parsePlcBtePapx should return error when data is out of bounds")
	}
	if err.Error() != "PlcBtePapx data out of bounds" {
		t.Errorf("unexpected error message: %v", err)
	}
}
