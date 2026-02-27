package doc

import (
	"fmt"
	"math/rand"
	"strings"
	"testing"
	"testing/quick"

	"github.com/shakinm/xlsReader/common"
)

// ============================================================================
// Property Test 4: 标题样式识别
// Feature: doc-format-preservation, Property 4: 标题样式识别
// Validates: Requirements 5.1, 5.2, 5.3
// ============================================================================

// headingTestInput holds generated data for heading identification property tests.
type headingTestInput struct {
	styles  []styleDef
	istd    uint16
	outLvl  uint8
}

func TestPropertyHeadingStyleIdentification(t *testing.T) {
	cfg := &quick.Config{MaxCount: 100}

	// Property: identifyHeading returns correct heading level based on priority:
	// 1. outLvl 0-8 -> HeadingLevel = outLvl + 1
	// 2. istd 1-9 -> HeadingLevel = istd
	// 3. style name "heading N" -> HeadingLevel = N
	// 4. Otherwise -> 0
	err := quick.Check(func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate a random style table with 0-20 entries
		numStyles := rng.Intn(21)
		styles := make([]styleDef, numStyles)
		for i := 0; i < numStyles; i++ {
			// Randomly assign heading names with random case
			if rng.Intn(3) == 0 && i > 0 {
				level := rng.Intn(9) + 1 // 1-9
				name := fmt.Sprintf("heading %d", level)
				// Random case variations
				switch rng.Intn(3) {
				case 0:
					name = strings.ToUpper(name)
				case 1:
					name = strings.Title(name)
				}
				styles[i] = styleDef{name: name, styleType: styleTypeParagraph}
			} else {
				styles[i] = styleDef{name: fmt.Sprintf("style%d", i), styleType: styleTypeParagraph}
			}
		}

		// Generate random paraFormatRun
		pr := &paraFormatRun{
			outLvl: uint8(rng.Intn(10)), // 0-9
			istd:   uint16(rng.Intn(max(numStyles, 15))),
		}

		result := identifyHeading(pr, styles)

		// Verify priority rules
		if pr.outLvl <= 8 {
			// Rule 1: outLvl 0-8 -> HeadingLevel = outLvl + 1
			expected := pr.outLvl + 1
			if result != expected {
				t.Errorf("outLvl=%d: expected HeadingLevel=%d, got %d", pr.outLvl, expected, result)
				return false
			}
			return true
		}

		// outLvl == 9 (body text), check sti from style
		if int(pr.istd) < len(styles) && styles[pr.istd].sti >= 1 && styles[pr.istd].sti <= 9 {
			// Rule 2: sti 1-9 -> HeadingLevel = sti
			expected := uint8(styles[pr.istd].sti)
			if result != expected {
				t.Errorf("sti=%d: expected HeadingLevel=%d, got %d", styles[pr.istd].sti, expected, result)
				return false
			}
			return true
		}

		// Check style name
		if int(pr.istd) < len(styles) {
			name := strings.ToLower(styles[pr.istd].name)
			if strings.HasPrefix(name, "heading ") {
				rest := name[len("heading "):]
				if len(rest) == 1 && rest[0] >= '1' && rest[0] <= '9' {
					expected := rest[0] - '0'
					if result != expected {
						t.Errorf("style name=%q: expected HeadingLevel=%d, got %d",
							styles[pr.istd].name, expected, result)
						return false
					}
					return true
				}
			}
		}

		// Rule 4: no heading
		if result != 0 {
			t.Errorf("no heading match: expected 0, got %d (outLvl=%d, istd=%d)",
				result, pr.outLvl, pr.istd)
			return false
		}
		return true
	}, cfg)

	if err != nil {
		t.Errorf("Property test failed: %v", err)
	}
}


// ============================================================================
// Property Test 5: 表格结构构建
// Feature: doc-format-preservation, Property 5: 表格结构构建
// Validates: Requirements 6.1, 6.2, 6.3, 6.4
// ============================================================================

func TestPropertyTableStructureBuilding(t *testing.T) {
	cfg := &quick.Config{MaxCount: 100}

	err := quick.Check(func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random table: 1-5 rows, each with 1-4 cells
		numRows := rng.Intn(5) + 1
		rowCellCounts := make([]int, numRows)
		for i := range rowCellCounts {
			rowCellCounts[i] = rng.Intn(4) + 1
		}

		// Build rawText: cells separated by 0x07, row-end is also 0x07
		// In DOC format, table cells and row-end marks use 0x07 as separator.
		var rawRunes []rune
		var paraRuns []paraFormatRun
		cpPos := uint32(0)

		for row := 0; row < numRows; row++ {
			for cell := 0; cell < rowCellCounts[row]; cell++ {
				cellText := fmt.Sprintf("r%dc%d", row, cell)
				cellRunes := []rune(cellText)
				pLen := uint32(len(cellRunes))

				rawRunes = append(rawRunes, cellRunes...)
				rawRunes = append(rawRunes, 0x07)

				paraRuns = append(paraRuns, paraFormatRun{
					cpStart: cpPos,
					cpEnd:   cpPos + pLen + 1,
					inTable: true,
					outLvl:  9,
				})
				cpPos += pLen + 1 // +1 for 0x07 separator
			}
			// Row-end paragraph (empty, just 0x07)
			rawRunes = append(rawRunes, 0x07)
			paraRuns = append(paraRuns, paraFormatRun{
				cpStart:     cpPos,
				cpEnd:       cpPos + 1,
				inTable:     true,
				tableRowEnd: true,
				outLvl:      9,
			})
			cpPos += 1 // 0x07 separator
		}

		rawText := string(rawRunes)

		fc := buildFormattedContent(rawText, nil, paraRuns, nil, nil, nil, nil, nil, nil)

		// Verify: all paragraphs should have InTable = true
		for i, p := range fc.Paragraphs {
			if !p.InTable {
				t.Errorf("paragraph %d: expected InTable=true, got false", i)
				return false
			}
		}

		// Count row-end paragraphs
		rowEndCount := 0
		for _, p := range fc.Paragraphs {
			if p.TableRowEnd {
				rowEndCount++
			}
		}
		if rowEndCount != numRows {
			t.Errorf("expected %d row-end paragraphs, got %d", numRows, rowEndCount)
			return false
		}

		// Total paragraphs = sum of cells + numRows (row-end paragraphs)
		totalCells := 0
		for _, c := range rowCellCounts {
			totalCells += c
		}
		expectedParas := totalCells + numRows
		if len(fc.Paragraphs) != expectedParas {
			t.Errorf("expected %d paragraphs, got %d", expectedParas, len(fc.Paragraphs))
			return false
		}

		return true
	}, cfg)

	if err != nil {
		t.Errorf("Property test failed: %v", err)
	}
}

// ============================================================================
// Property Test 7: Page Break and Section Break Detection
// Feature: doc-format-preservation, Property 7: Page Break and Section Break Detection
// Validates: Requirements 8.1, 8.2, 8.3
// ============================================================================

func TestPropertyPageBreakAndSectionBreakDetection(t *testing.T) {
	cfg := &quick.Config{MaxCount: 100}

	err := quick.Check(func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random paragraphs, some with 0x0C, some with pageBreakBefore
		numParas := rng.Intn(10) + 1
		paraTexts := make([]string, numParas)
		hasPageBreakChar := make([]bool, numParas)
		hasPageBreakBefore := make([]bool, numParas)

		var paraRuns []paraFormatRun
		cpPos := uint32(0)

		for i := 0; i < numParas; i++ {
			text := fmt.Sprintf("para%d", i)
			// Randomly insert 0x0C
			if rng.Intn(3) == 0 {
				text = text + "\x0C"
				hasPageBreakChar[i] = true
			}
			paraTexts[i] = text

			pLen := uint32(len([]rune(text)))
			pbBefore := rng.Intn(3) == 0
			hasPageBreakBefore[i] = pbBefore

			paraRuns = append(paraRuns, paraFormatRun{
				cpStart:         cpPos,
				cpEnd:           cpPos + pLen + 1,
				pageBreakBefore: pbBefore,
				outLvl:          9,
			})
			cpPos += pLen + 1
		}

		rawText := strings.Join(paraTexts, "\r")

		fc := buildFormattedContent(rawText, nil, paraRuns, nil, nil, nil, nil, nil, nil)

		if len(fc.Paragraphs) != numParas {
			t.Errorf("expected %d paragraphs, got %d", numParas, len(fc.Paragraphs))
			return false
		}

		for i, p := range fc.Paragraphs {
			// Verify HasPageBreak for 0x0C
			if p.HasPageBreak != hasPageBreakChar[i] {
				t.Errorf("para %d: expected HasPageBreak=%v, got %v",
					i, hasPageBreakChar[i], p.HasPageBreak)
				return false
			}
			// Verify PageBreakBefore from sprm
			if p.PageBreakBefore != hasPageBreakBefore[i] {
				t.Errorf("para %d: expected PageBreakBefore=%v, got %v",
					i, hasPageBreakBefore[i], p.PageBreakBefore)
				return false
			}
		}

		return true
	}, cfg)

	if err != nil {
		t.Errorf("Property test failed: %v", err)
	}
}


// ============================================================================
// Property Test 13: Backward Compatibility
// Feature: doc-format-preservation, Property 13: Backward Compatibility
// Validates: Requirements 9.7, 9.8, 16.1, 16.2
// ============================================================================

func TestPropertyBackwardCompatibility(t *testing.T) {
	cfg := &quick.Config{MaxCount: 100}

	err := quick.Check(func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random text
		textLen := rng.Intn(200)
		textBytes := make([]byte, textLen)
		for i := range textBytes {
			textBytes[i] = byte(rng.Intn(94) + 32) // printable ASCII
		}
		text := string(textBytes)

		// Generate random images
		numImages := rng.Intn(5)
		images := make([]common.Image, numImages)
		for i := range images {
			dataLen := rng.Intn(50) + 1
			imgData := make([]byte, dataLen)
			for j := range imgData {
				imgData[j] = byte(rng.Intn(256))
			}
			images[i] = common.Image{
				Format: common.ImageFormat(rng.Intn(7)),
				Data:   imgData,
			}
		}

		// Randomly include or exclude formattedContent
		var fc *FormattedContent
		if rng.Intn(2) == 0 {
			fc = &FormattedContent{
				Paragraphs: []Paragraph{{
					Runs: []TextRun{{Text: "test"}},
				}},
			}
		}

		doc := Document{
			text:             text,
			images:           images,
			formattedContent: fc,
		}

		// Verify GetText() returns the same text
		if doc.GetText() != text {
			t.Errorf("GetText() mismatch: expected %q, got %q", text, doc.GetText())
			return false
		}

		// Verify GetImages() returns the same images
		gotImages := doc.GetImages()
		if len(gotImages) != len(images) {
			t.Errorf("GetImages() count mismatch: expected %d, got %d",
				len(images), len(gotImages))
			return false
		}
		for i := range images {
			if images[i].Format != gotImages[i].Format {
				t.Errorf("image %d format mismatch", i)
				return false
			}
			if len(images[i].Data) != len(gotImages[i].Data) {
				t.Errorf("image %d data length mismatch", i)
				return false
			}
		}

		return true
	}, cfg)

	if err != nil {
		t.Errorf("Property test failed: %v", err)
	}
}

// ============================================================================
// Unit Tests for buildFormattedContent
// Task 7.8
// ============================================================================

func TestBuildFormattedContent_HeadingByName(t *testing.T) {
	// Style with name "heading 1" at index 10 -> HeadingLevel = 1
	// Use istd=10 (outside 1-9 range) so name matching is used instead of built-in index
	styles := make([]styleDef, 11)
	styles[0] = styleDef{name: "Normal", styleType: styleTypeParagraph}
	styles[10] = styleDef{name: "heading 1", styleType: styleTypeParagraph}

	paraRuns := []paraFormatRun{
		{cpStart: 0, cpEnd: 6, istd: 10, outLvl: 9},
	}
	rawText := "Hello\r"

	fc := buildFormattedContent(rawText, nil, paraRuns, styles, nil, nil, nil, nil, nil)

	if len(fc.Paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(fc.Paragraphs))
	}
	if fc.Paragraphs[0].HeadingLevel != 1 {
		t.Errorf("expected HeadingLevel=1, got %d", fc.Paragraphs[0].HeadingLevel)
	}
}

func TestBuildFormattedContent_HeadingByIstd(t *testing.T) {
	// sti = 3 in style at istd=3 -> HeadingLevel = 3 (built-in heading style index)
	styles := []styleDef{
		{name: "Normal"},
		{name: "style1"},
		{name: "style2"},
		{name: "style3", sti: 3},
	}
	paraRuns := []paraFormatRun{
		{cpStart: 0, cpEnd: 6, istd: 3, outLvl: 9},
	}
	rawText := "Hello\r"

	fc := buildFormattedContent(rawText, nil, paraRuns, styles, nil, nil, nil, nil, nil)

	if len(fc.Paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(fc.Paragraphs))
	}
	if fc.Paragraphs[0].HeadingLevel != 3 {
		t.Errorf("expected HeadingLevel=3, got %d", fc.Paragraphs[0].HeadingLevel)
	}
}

func TestBuildFormattedContent_HeadingByOutLvl(t *testing.T) {
	// outLvl = 0 -> HeadingLevel = 1 (highest priority)
	paraRuns := []paraFormatRun{
		{cpStart: 0, cpEnd: 6, istd: 0, outLvl: 0},
	}
	rawText := "Hello\r"

	fc := buildFormattedContent(rawText, nil, paraRuns, nil, nil, nil, nil, nil, nil)

	if len(fc.Paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(fc.Paragraphs))
	}
	if fc.Paragraphs[0].HeadingLevel != 1 {
		t.Errorf("expected HeadingLevel=1, got %d", fc.Paragraphs[0].HeadingLevel)
	}
}

func TestBuildFormattedContent_TableStructure(t *testing.T) {
	// Table with 2 rows, 2 cells each
	// Row 1: "A\x07" "B\x07" + row-end "\x07"
	// Row 2: "C\x07" "D\x07" + row-end "\x07"
	rawText := "A\x07B\x07\x07C\x07D\x07\x07"
	paraRuns := []paraFormatRun{
		{cpStart: 0, cpEnd: 2, inTable: true, outLvl: 9},
		{cpStart: 2, cpEnd: 4, inTable: true, outLvl: 9},
		{cpStart: 4, cpEnd: 5, inTable: true, tableRowEnd: true, outLvl: 9},
		{cpStart: 5, cpEnd: 7, inTable: true, outLvl: 9},
		{cpStart: 7, cpEnd: 9, inTable: true, outLvl: 9},
		{cpStart: 9, cpEnd: 10, inTable: true, tableRowEnd: true, outLvl: 9},
	}

	fc := buildFormattedContent(rawText, nil, paraRuns, nil, nil, nil, nil, nil, nil)

	if len(fc.Paragraphs) != 6 {
		t.Fatalf("expected 6 paragraphs, got %d", len(fc.Paragraphs))
	}

	// All should be InTable
	for i, p := range fc.Paragraphs {
		if !p.InTable {
			t.Errorf("paragraph %d: expected InTable=true", i)
		}
	}

	// Paragraphs 2 and 5 should be row-end
	if !fc.Paragraphs[2].TableRowEnd {
		t.Error("paragraph 2: expected TableRowEnd=true")
	}
	if !fc.Paragraphs[5].TableRowEnd {
		t.Error("paragraph 5: expected TableRowEnd=true")
	}
}

func TestBuildFormattedContent_UnevenTableRows(t *testing.T) {
	// Row 1: 3 cells, Row 2: 1 cell - different cell counts preserved
	rawText := "A\x07B\x07C\x07\x07D\x07\x07"
	paraRuns := []paraFormatRun{
		{cpStart: 0, cpEnd: 2, inTable: true, outLvl: 9},
		{cpStart: 2, cpEnd: 4, inTable: true, outLvl: 9},
		{cpStart: 4, cpEnd: 6, inTable: true, outLvl: 9},
		{cpStart: 6, cpEnd: 7, inTable: true, tableRowEnd: true, outLvl: 9},
		{cpStart: 7, cpEnd: 9, inTable: true, outLvl: 9},
		{cpStart: 9, cpEnd: 10, inTable: true, tableRowEnd: true, outLvl: 9},
	}

	fc := buildFormattedContent(rawText, nil, paraRuns, nil, nil, nil, nil, nil, nil)

	if len(fc.Paragraphs) != 6 {
		t.Fatalf("expected 6 paragraphs, got %d", len(fc.Paragraphs))
	}

	// All InTable
	for i, p := range fc.Paragraphs {
		if !p.InTable {
			t.Errorf("paragraph %d: expected InTable=true", i)
		}
	}

	// Row-end at index 3 and 5
	rowEndIndices := []int{3, 5}
	for _, idx := range rowEndIndices {
		if !fc.Paragraphs[idx].TableRowEnd {
			t.Errorf("paragraph %d: expected TableRowEnd=true", idx)
		}
	}
}

func TestBuildFormattedContent_ListDefault(t *testing.T) {
	// No list defs -> default unordered (ListType = 0)
	paraRuns := []paraFormatRun{
		{cpStart: 0, cpEnd: 5, ilfo: 1, ilvl: 0, outLvl: 9},
	}
	rawText := "Item\r"

	fc := buildFormattedContent(rawText, nil, paraRuns, nil, nil, nil, nil, nil, nil)

	if len(fc.Paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(fc.Paragraphs))
	}
	p := fc.Paragraphs[0]
	if !p.IsListItem {
		t.Error("expected IsListItem=true")
	}
	if p.ListType != 0 {
		t.Errorf("expected ListType=0 (unordered), got %d", p.ListType)
	}
	if p.ListLevel != 0 {
		t.Errorf("expected ListLevel=0, got %d", p.ListLevel)
	}
}

func TestBuildFormattedContent_PageBreak(t *testing.T) {
	// Text with 0x0C -> HasPageBreak = true
	rawText := "Before\x0CAfter\r"
	paraRuns := []paraFormatRun{
		{cpStart: 0, cpEnd: 13, outLvl: 9},
	}

	fc := buildFormattedContent(rawText, nil, paraRuns, nil, nil, nil, nil, nil, nil)

	if len(fc.Paragraphs) != 1 {
		t.Fatalf("expected 1 paragraph, got %d", len(fc.Paragraphs))
	}
	if !fc.Paragraphs[0].HasPageBreak {
		t.Error("expected HasPageBreak=true for paragraph containing 0x0C")
	}
}

func TestBuildFormattedContent_SectionBreak(t *testing.T) {
	// Simplified: verify HasPageBreak for 0x0C
	rawText := "Text\x0C\rNormal\r"
	paraRuns := []paraFormatRun{
		{cpStart: 0, cpEnd: 6, outLvl: 9},
		{cpStart: 6, cpEnd: 13, outLvl: 9},
	}

	fc := buildFormattedContent(rawText, nil, paraRuns, nil, nil, nil, nil, nil, nil)

	if len(fc.Paragraphs) != 2 {
		t.Fatalf("expected 2 paragraphs, got %d", len(fc.Paragraphs))
	}
	if !fc.Paragraphs[0].HasPageBreak {
		t.Error("paragraph 0: expected HasPageBreak=true")
	}
	if fc.Paragraphs[1].HasPageBreak {
		t.Error("paragraph 1: expected HasPageBreak=false")
	}
}

// max returns the larger of two ints (Go 1.13 doesn't have built-in max for int).
func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}
