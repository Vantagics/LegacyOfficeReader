package docconv

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math/rand"
	"strings"
	"testing"
	"testing/quick"

	"github.com/shakinm/xlsReader/common"
	"github.com/shakinm/xlsReader/doc"
)

func TestConvertReader_InvalidInput(t *testing.T) {
	// Provide invalid (non-DOC) data
	input := bytes.NewReader([]byte("this is not a valid DOC file"))
	var output bytes.Buffer

	err := ConvertReader(input, &output)
	if err == nil {
		t.Fatal("expected error for invalid input, got nil")
	}
	if !strings.Contains(err.Error(), "docconv") {
		t.Errorf("error message should contain 'docconv' prefix, got: %s", err.Error())
	}
}

func TestConvertFile_NonexistentInput(t *testing.T) {
	err := ConvertFile("/nonexistent/path/to/file.doc", "/tmp/output.docx")
	if err == nil {
		t.Fatal("expected error for nonexistent input file, got nil")
	}
	if !strings.Contains(err.Error(), "docconv") {
		t.Errorf("error message should contain 'docconv' prefix, got: %s", err.Error())
	}
}

func TestConvertFile_InvalidOutputPath(t *testing.T) {
	err := ConvertFile("/nonexistent/input.doc", "/nonexistent/dir/output.docx")
	if err == nil {
		t.Fatal("expected error, got nil")
	}
	if !strings.Contains(err.Error(), "docconv") {
		t.Errorf("error message should contain 'docconv' prefix, got: %s", err.Error())
	}
}

// docRandomString generates a random alphanumeric string of the given length.
func docRandomString(rng *rand.Rand, length int) string {
	const charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 "
	b := make([]byte, length)
	for i := range b {
		b[i] = charset[rng.Intn(len(charset))]
	}
	return string(b)
}

// docReadZipFile reads the content of a file inside a zip archive.
func docReadZipFile(zr *zip.Reader, name string) (string, error) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return "", err
			}
			defer rc.Close()
			var buf bytes.Buffer
			if _, err := io.Copy(&buf, rc); err != nil {
				return "", err
			}
			return buf.String(), nil
		}
	}
	return "", fmt.Errorf("file %s not found in zip", name)
}

// docEscapeXML escapes a string the same way xml.Escape does.
func docEscapeXML(s string) string {
	var buf bytes.Buffer
	xml.Escape(&buf, []byte(s))
	return buf.String()
}

// Feature: legacy-to-ooxml-conversion, Property 5: DOC 文本保留
// **Validates: Requirements 3.3**
//
// For any text string, after mapping through writeDocx, the output DOCX
// word/document.xml should contain the text (XML-escaped) exactly.
func TestProperty_DocTextPreservation(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random text (1-200 chars, alphanumeric + spaces)
		textLen := 1 + rng.Intn(200)
		text := docRandomString(rng, textLen)

		// Create textData with that text and empty images
		td := textData{Text: text}

		// Call writeDocx to a buffer
		var buf bytes.Buffer
		if err := writeDocx(&buf, td, nil); err != nil {
			t.Logf("writeDocx failed: %v", err)
			return false
		}

		// Open the zip, read word/document.xml
		reader := bytes.NewReader(buf.Bytes())
		zr, err := zip.NewReader(reader, int64(buf.Len()))
		if err != nil {
			t.Logf("failed to open zip: %v", err)
			return false
		}

		content, err := docReadZipFile(zr, "word/document.xml")
		if err != nil {
			t.Logf("failed to read word/document.xml: %v", err)
			return false
		}

		// Verify the XML-escaped text appears in the document
		escaped := docEscapeXML(text)
		if !strings.Contains(content, escaped) {
			t.Logf("document.xml missing text %q (escaped: %q)", text, escaped)
			return false
		}

		// Also verify by parsing the XML to extract <w:t> content
		type WText struct {
			Value string `xml:",chardata"`
		}
		type WRun struct {
			Text WText `xml:"t"`
		}
		type WParagraph struct {
			Runs []WRun `xml:"r"`
		}
		type WBody struct {
			Paragraphs []WParagraph `xml:"p"`
		}
		type WDocument struct {
			Body WBody `xml:"body"`
		}

		var doc WDocument
		if err := xml.Unmarshal([]byte(content), &doc); err != nil {
			t.Logf("failed to parse document.xml: %v", err)
			return false
		}

		// Collect all <w:t> text content
		var allText strings.Builder
		for _, p := range doc.Body.Paragraphs {
			for _, r := range p.Runs {
				allText.WriteString(r.Text.Value)
			}
		}

		if allText.String() != text {
			t.Logf("parsed text %q does not match input %q", allText.String(), text)
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: DOC text preservation: %v", err)
	}
}

// docReadZipFileBytes reads the raw bytes of a file inside a zip archive.
func docReadZipFileBytes(zr *zip.Reader, name string) ([]byte, error) {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			var buf bytes.Buffer
			if _, err := io.Copy(&buf, rc); err != nil {
				return nil, err
			}
			return buf.Bytes(), nil
		}
	}
	return nil, fmt.Errorf("file %s not found in zip", name)
}

// Feature: legacy-to-ooxml-conversion, Property 6: DOC 图片保留
// **Validates: Requirements 3.4**
//
// For any set of images with arbitrary format and data, after mapping through writeDocx,
// the output DOCX should contain all images with matching data in word/media/.
func TestProperty_DocImagePreservation(t *testing.T) {
	allFormats := []common.ImageFormat{
		common.ImageFormatEMF,
		common.ImageFormatWMF,
		common.ImageFormatPICT,
		common.ImageFormatJPEG,
		common.ImageFormatPNG,
		common.ImageFormatDIB,
		common.ImageFormatTIFF,
	}

	config := &quick.Config{MaxCount: 100}

	prop := func(imageCount uint8, seed int64) bool {
		// Limit image count to a reasonable range (0-8)
		numImages := int(imageCount) % 9
		rng := rand.New(rand.NewSource(seed))

		// Generate random images
		images := make([]imageData, numImages)
		for i := 0; i < numImages; i++ {
			format := allFormats[rng.Intn(len(allFormats))]
			dataLen := 1 + rng.Intn(100)
			data := make([]byte, dataLen)
			for j := range data {
				data[j] = byte(rng.Intn(256))
			}
			images[i] = imageData{Format: format, Data: data}
		}

		// Create empty textData and call writeDocx
		td := textData{Text: ""}
		var buf bytes.Buffer
		if err := writeDocx(&buf, td, images); err != nil {
			t.Logf("writeDocx failed: %v", err)
			return false
		}

		// Open the zip
		reader := bytes.NewReader(buf.Bytes())
		zr, err := zip.NewReader(reader, int64(buf.Len()))
		if err != nil {
			t.Logf("failed to open zip: %v", err)
			return false
		}

		// Count image files in word/media/
		imageFiles := 0
		for _, f := range zr.File {
			if strings.HasPrefix(f.Name, "word/media/") {
				imageFiles++
			}
		}

		if imageFiles != numImages {
			t.Logf("expected %d images in word/media/, got %d", numImages, imageFiles)
			return false
		}

		// Verify each image's data matches
		for i, img := range images {
			ext := (&common.Image{Format: img.Format}).Extension()
			if ext == "" {
				ext = ".bin"
			}
			filename := fmt.Sprintf("word/media/image%d%s", i+1, ext)
			content, err := docReadZipFileBytes(zr, filename)
			if err != nil {
				t.Logf("failed to read %s: %v", filename, err)
				return false
			}
			if !bytes.Equal(content, img.Data) {
				t.Logf("image %d data mismatch", i+1)
				return false
			}
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: DOC image preservation: %v", err)
	}
}

// Feature: doc-format-preservation, Property 8: 字符格式往返一致性
// **Validates: Requirements 17.1, 10.1, 10.2, 10.3, 10.4, 10.5, 10.6**
//
// For any CharacterFormatting, writing it to DOCX <w:rPr> XML and parsing
// the output should yield semantically equivalent formatting.
func TestProperty_CharFormatRoundTrip(t *testing.T) {
	validFonts := []string{"Arial", "Times New Roman", "Calibri", "Courier New", "Verdana"}

	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random CharacterFormatting
		fontName := validFonts[rng.Intn(len(validFonts))]
		fontSize := uint16(1 + rng.Intn(100))
		bold := rng.Intn(2) == 1
		italic := rng.Intn(2) == 1
		underline := uint8(rng.Intn(6))
		color := fmt.Sprintf("%06X", rng.Intn(0xFFFFFF+1))

		props := doc.CharacterFormatting{
			FontName:  fontName,
			FontSize:  fontSize,
			Bold:      bold,
			Italic:    italic,
			Underline: underline,
			Color:     color,
		}

		// Create formattedData with one paragraph, one run
		fd := &formattedData{
			Paragraphs: []formattedParagraph{
				{
					Runs: []formattedRun{
						{Text: "test text", Props: props},
					},
				},
			},
		}

		// Write to buffer
		var buf bytes.Buffer
		if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
			t.Logf("writeFormattedDocumentXML failed: %v", err)
			return false
		}
		output := buf.String()

		// Verify FontName
		expectedFont := fmt.Sprintf(`<w:rFonts w:ascii="%s"`, fontName)
		if !strings.Contains(output, expectedFont) {
			t.Logf("missing font: expected %s in output", expectedFont)
			return false
		}

		// Verify FontSize
		expectedSize := fmt.Sprintf(`<w:sz w:val="%d"/>`, fontSize)
		if !strings.Contains(output, expectedSize) {
			t.Logf("missing font size: expected %s", expectedSize)
			return false
		}

		// Verify Bold
		if bold {
			if !strings.Contains(output, `<w:b/>`) {
				t.Logf("missing <w:b/> for bold=true")
				return false
			}
		} else {
			if strings.Contains(output, `<w:b/>`) {
				t.Logf("unexpected <w:b/> for bold=false")
				return false
			}
		}

		// Verify Italic
		if italic {
			if !strings.Contains(output, `<w:i/>`) {
				t.Logf("missing <w:i/> for italic=true")
				return false
			}
		} else {
			if strings.Contains(output, `<w:i/>`) {
				t.Logf("unexpected <w:i/> for italic=false")
				return false
			}
		}

		// Verify Underline
		if underline > 0 {
			if !strings.Contains(output, `<w:u`) {
				t.Logf("missing <w:u> for underline=%d", underline)
				return false
			}
		} else {
			if strings.Contains(output, `<w:u`) {
				t.Logf("unexpected <w:u> for underline=0")
				return false
			}
		}

		// Verify Color
		expectedColor := fmt.Sprintf(`<w:color w:val="%s"/>`, color)
		if !strings.Contains(output, expectedColor) {
			t.Logf("missing color: expected %s", expectedColor)
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: character format round-trip: %v", err)
	}
}

// Feature: doc-format-preservation, Property 9: 段落格式往返一致性
// **Validates: Requirements 17.2, 17.3, 11.1, 11.2, 11.3, 12.1**
//
// For any ParagraphFormatting, writing it to DOCX <w:pPr> XML and parsing
// the output should yield semantically equivalent formatting.
func TestProperty_ParaFormatRoundTrip(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random ParagraphFormatting
		alignment := uint8(rng.Intn(4))
		indentLeft := int32(rng.Intn(10000))
		indentRight := int32(rng.Intn(10000))
		indentFirst := int32(rng.Intn(10000))
		spaceBefore := uint16(rng.Intn(5000))
		spaceAfter := uint16(rng.Intn(5000))
		lineSpacing := int32(rng.Intn(5000))

		props := doc.ParagraphFormatting{
			Alignment:   alignment,
			IndentLeft:  indentLeft,
			IndentRight: indentRight,
			IndentFirst: indentFirst,
			SpaceBefore: spaceBefore,
			SpaceAfter:  spaceAfter,
			LineSpacing: lineSpacing,
		}

		// Create formattedData with one paragraph
		fd := &formattedData{
			Paragraphs: []formattedParagraph{
				{
					Props: props,
					Runs: []formattedRun{
						{Text: "test", Props: doc.CharacterFormatting{}},
					},
				},
			},
		}

		// Write to buffer
		var buf bytes.Buffer
		if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
			t.Logf("writeFormattedDocumentXML failed: %v", err)
			return false
		}
		output := buf.String()

		// Verify Alignment
		switch alignment {
		case 0:
			// No <w:jc> element for left alignment
			if strings.Contains(output, `<w:jc`) {
				t.Logf("unexpected <w:jc> for alignment=0 (left)")
				return false
			}
		case 1:
			if !strings.Contains(output, `<w:jc w:val="center"/>`) {
				t.Logf("missing center alignment")
				return false
			}
		case 2:
			if !strings.Contains(output, `<w:jc w:val="right"/>`) {
				t.Logf("missing right alignment")
				return false
			}
		case 3:
			if !strings.Contains(output, `<w:jc w:val="both"/>`) {
				t.Logf("missing both alignment")
				return false
			}
		}

		// Verify Indentation
		if indentLeft != 0 || indentRight != 0 || indentFirst != 0 {
			if !strings.Contains(output, `<w:ind`) {
				t.Logf("missing <w:ind> element")
				return false
			}
			if indentLeft != 0 {
				expected := fmt.Sprintf(`w:left="%d"`, indentLeft)
				if !strings.Contains(output, expected) {
					t.Logf("missing indent left: %s", expected)
					return false
				}
			}
			if indentRight != 0 {
				expected := fmt.Sprintf(`w:right="%d"`, indentRight)
				if !strings.Contains(output, expected) {
					t.Logf("missing indent right: %s", expected)
					return false
				}
			}
			if indentFirst > 0 {
				expected := fmt.Sprintf(`w:firstLine="%d"`, indentFirst)
				if !strings.Contains(output, expected) {
					t.Logf("missing indent first: %s", expected)
					return false
				}
			}
		}

		// Verify Spacing
		if spaceBefore != 0 || spaceAfter != 0 || lineSpacing != 0 {
			if !strings.Contains(output, `<w:spacing`) {
				t.Logf("missing <w:spacing> element")
				return false
			}
			if spaceBefore != 0 {
				expected := fmt.Sprintf(`w:before="%d"`, spaceBefore)
				if !strings.Contains(output, expected) {
					t.Logf("missing space before: %s", expected)
					return false
				}
			}
			if spaceAfter != 0 {
				expected := fmt.Sprintf(`w:after="%d"`, spaceAfter)
				if !strings.Contains(output, expected) {
					t.Logf("missing space after: %s", expected)
					return false
				}
			}
			if lineSpacing != 0 {
				expected := fmt.Sprintf(`w:line="%d"`, lineSpacing)
				if !strings.Contains(output, expected) {
					t.Logf("missing line spacing: %s", expected)
					return false
				}
			}
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: paragraph format round-trip: %v", err)
	}
}

// Feature: doc-format-preservation, Property 12: 分页与OOXML 输出
// **Validates: Requirements 15.1, 15.2**
//
// For any paragraphs with PageBreakBefore and HasPageBreak flags,
// the generated XML should contain the corresponding elements.
func TestProperty_PageBreakOutput(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random paragraphs
		numParas := 1 + rng.Intn(5)
		paragraphs := make([]formattedParagraph, numParas)
		for i := 0; i < numParas; i++ {
			paragraphs[i] = formattedParagraph{
				PageBreakBefore: rng.Intn(2) == 1,
				HasPageBreak:    rng.Intn(2) == 1,
				Runs: []formattedRun{
					{Text: fmt.Sprintf("para%d", i), Props: doc.CharacterFormatting{}},
				},
			}
		}

		fd := &formattedData{Paragraphs: paragraphs}

		var buf bytes.Buffer
		if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
			t.Logf("writeFormattedDocumentXML failed: %v", err)
			return false
		}
		output := buf.String()

		// Count occurrences
		pageBreakBeforeCount := strings.Count(output, `<w:pageBreakBefore/>`)
		brPageCount := strings.Count(output, `<w:br w:type="page"/>`)

		expectedPBB := 0
		expectedBR := 0
		for _, p := range paragraphs {
			if p.PageBreakBefore {
				expectedPBB++
			}
			if p.HasPageBreak {
				expectedBR++
			}
		}

		if pageBreakBeforeCount != expectedPBB {
			t.Logf("pageBreakBefore count: got %d, want %d", pageBreakBeforeCount, expectedPBB)
			return false
		}
		if brPageCount != expectedBR {
			t.Logf("br page count: got %d, want %d", brPageCount, expectedBR)
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: page break OOXML output: %v", err)
	}
}

// TestWriteFormattedDocumentXML_CharFormat verifies character formatting XML elements.
func TestWriteFormattedDocumentXML_CharFormat(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{
				Runs: []formattedRun{
					{
						Text: "Hello",
						Props: doc.CharacterFormatting{
							FontName:  "Arial",
							FontSize:  24,
							Bold:      true,
							Italic:    true,
							Underline: 1,
							Color:     "FF0000",
						},
					},
				},
			},
		},
	}

	var buf bytes.Buffer
	if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
		t.Fatalf("writeFormattedDocumentXML failed: %v", err)
	}
	output := buf.String()

	checks := []struct {
		name     string
		expected string
	}{
		{"font", `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`},
		{"bold", `<w:b/>`},
		{"italic", `<w:i/>`},
		{"underline", `<w:u w:val="single"/>`},
		{"fontSize", `<w:sz w:val="24"/>`},
		{"color", `<w:color w:val="FF0000"/>`},
		{"text", `Hello`},
	}

	for _, c := range checks {
		if !strings.Contains(output, c.expected) {
			t.Errorf("missing %s: expected %q in output", c.name, c.expected)
		}
	}
}

// TestWriteFormattedDocumentXML_ParaFormat verifies paragraph formatting XML elements.
func TestWriteFormattedDocumentXML_ParaFormat(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{
				Props: doc.ParagraphFormatting{
					Alignment:   1, // center
					IndentLeft:  720,
					IndentRight: 360,
					IndentFirst: 480,
					SpaceBefore: 120,
					SpaceAfter:  240,
					LineSpacing:  360,
				},
				Runs: []formattedRun{
					{Text: "Centered", Props: doc.CharacterFormatting{}},
				},
			},
		},
	}

	var buf bytes.Buffer
	if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
		t.Fatalf("writeFormattedDocumentXML failed: %v", err)
	}
	output := buf.String()

	checks := []struct {
		name     string
		expected string
	}{
		{"alignment", `<w:jc w:val="center"/>`},
		{"indentLeft", `w:left="720"`},
		{"indentRight", `w:right="360"`},
		{"indentFirst", `w:firstLine="480"`},
		{"spaceBefore", `w:before="120"`},
		{"spaceAfter", `w:after="240"`},
		{"lineSpacing", `w:line="360"`},
	}

	for _, c := range checks {
		if !strings.Contains(output, c.expected) {
			t.Errorf("missing %s: expected %q in output", c.name, c.expected)
		}
	}
}

// TestWriteFormattedDocumentXML_PageBreak verifies page break XML elements.
func TestWriteFormattedDocumentXML_PageBreak(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{
				PageBreakBefore: true,
				HasPageBreak:    true,
				Runs: []formattedRun{
					{Text: "After break", Props: doc.CharacterFormatting{}},
				},
			},
		},
	}

	var buf bytes.Buffer
	if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
		t.Fatalf("writeFormattedDocumentXML failed: %v", err)
	}
	output := buf.String()

	if !strings.Contains(output, `<w:pageBreakBefore/>`) {
		t.Error("missing <w:pageBreakBefore/>")
	}
	if !strings.Contains(output, `<w:br w:type="page"/>`) {
		t.Error("missing <w:br w:type=\"page\"/>")
	}
}

// TestWriteDocx_PlainTextFallback verifies that writeDocx still works for plain text mode.
func TestWriteDocx_PlainTextFallback(t *testing.T) {
	td := textData{Text: "Hello World\nSecond paragraph"}
	var buf bytes.Buffer
	if err := writeDocx(&buf, td, nil); err != nil {
		t.Fatalf("writeDocx failed: %v", err)
	}

	// Verify it's a valid zip
	reader := bytes.NewReader(buf.Bytes())
	zr, err := zip.NewReader(reader, int64(buf.Len()))
	if err != nil {
		t.Fatalf("failed to open zip: %v", err)
	}

	// Verify word/document.xml exists and contains the text
	content, err := docReadZipFile(zr, "word/document.xml")
	if err != nil {
		t.Fatalf("failed to read word/document.xml: %v", err)
	}

	if !strings.Contains(content, "Hello World") {
		t.Error("document.xml missing 'Hello World'")
	}
	if !strings.Contains(content, "Second paragraph") {
		t.Error("document.xml missing 'Second paragraph'")
	}
}

// TestWriteFormattedDocumentXML_HeadingStyle verifies heading style reference is generated.
// _需求: 12.1, 12.3_
func TestWriteFormattedDocumentXML_HeadingStyle(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{
				HeadingLevel: 2,
				Runs: []formattedRun{
					{Text: "Heading Two", Props: doc.CharacterFormatting{}},
				},
			},
		},
	}

	var buf bytes.Buffer
	if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
		t.Fatalf("writeFormattedDocumentXML failed: %v", err)
	}
	output := buf.String()

	if !strings.Contains(output, `<w:pStyle w:val="Heading2"/>`) {
		t.Error("missing <w:pStyle w:val=\"Heading2\"/> for HeadingLevel=2")
	}
	// Verify pStyle is inside pPr
	if !strings.Contains(output, `<w:pPr><w:pStyle w:val="Heading2"/>`) {
		t.Error("<w:pStyle> should be the first element inside <w:pPr>")
	}
}

// TestWriteFormattedStylesXML_HeadingDefinitions verifies heading style definitions in styles.xml.
// _需求: 12.2_
func TestWriteFormattedStylesXML_HeadingDefinitions(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{
				HeadingLevel: 1,
				Runs:         []formattedRun{{Text: "H1", Props: doc.CharacterFormatting{}}},
			},
			{
				HeadingLevel: 3,
				Runs:         []formattedRun{{Text: "H3", Props: doc.CharacterFormatting{}}},
			},
			{
				HeadingLevel: 0, // not a heading
				Runs:         []formattedRun{{Text: "Normal", Props: doc.CharacterFormatting{}}},
			},
		},
	}

	var buf bytes.Buffer
	if err := writeFormattedStylesXML(&buf, fd); err != nil {
		t.Fatalf("writeFormattedStylesXML failed: %v", err)
	}
	output := buf.String()

	// Heading1 style definition
	if !strings.Contains(output, `<w:style w:type="paragraph" w:styleId="Heading1">`) {
		t.Error("missing Heading1 style definition")
	}
	if !strings.Contains(output, `<w:name w:val="heading 1"/>`) {
		t.Error("missing heading 1 name")
	}
	if !strings.Contains(output, `<w:outlineLvl w:val="0"/>`) {
		t.Error("missing outlineLvl 0 for Heading1")
	}

	// Heading3 style definition
	if !strings.Contains(output, `<w:style w:type="paragraph" w:styleId="Heading3">`) {
		t.Error("missing Heading3 style definition")
	}
	if !strings.Contains(output, `<w:name w:val="heading 3"/>`) {
		t.Error("missing heading 3 name")
	}
	if !strings.Contains(output, `<w:outlineLvl w:val="2"/>`) {
		t.Error("missing outlineLvl 2 for Heading3")
	}

	// Should NOT contain Heading2 (not used)
	if strings.Contains(output, `w:styleId="Heading2"`) {
		t.Error("unexpected Heading2 style definition (level 2 not used)")
	}

	// basedOn Normal
	if !strings.Contains(output, `<w:basedOn w:val="Normal"/>`) {
		t.Error("missing basedOn Normal")
	}
}

// TestWriteFormattedDocumentXML_NoHeading verifies no heading style reference when HeadingLevel=0.
// _需求: 12.3_
func TestWriteFormattedDocumentXML_NoHeading(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{
				HeadingLevel: 0,
				Runs: []formattedRun{
					{Text: "Normal paragraph", Props: doc.CharacterFormatting{}},
				},
			},
		},
	}

	var buf bytes.Buffer
	if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
		t.Fatalf("writeFormattedDocumentXML failed: %v", err)
	}
	output := buf.String()

	if strings.Contains(output, `<w:pStyle`) {
		t.Error("unexpected <w:pStyle> for HeadingLevel=0")
	}
}

// Feature: doc-format-preservation, Property 10: 表格 OOXML 输出
// **Validates: Requirements 13.1, 13.2, 13.3, 13.4**
//
// For any table structure (random row count and random cells per row),
// writeFormattedDocumentXML should produce XML with the correct number of
// <w:tbl>, <w:tr>, <w:tc> elements, and each <w:tbl> should contain
// <w:tblPr> and <w:tblBorders>.
func TestProperty_TableOOXMLOutput(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate random table structure: 1-5 rows, 1-5 cells per row
		numRows := 1 + rng.Intn(5)
		totalCells := 0
		var paragraphs []formattedParagraph

		for r := 0; r < numRows; r++ {
			numCells := 1 + rng.Intn(5)
			totalCells += numCells
			for c := 0; c < numCells; c++ {
				paragraphs = append(paragraphs, formattedParagraph{
					InTable:        true,
					IsTableCellEnd: true,
					Runs: []formattedRun{
						{Text: fmt.Sprintf("r%dc%d", r, c), Props: doc.CharacterFormatting{}},
					},
				})
			}
			// Row end marker
			paragraphs = append(paragraphs, formattedParagraph{
				InTable:     true,
				TableRowEnd: true,
			})
		}

		fd := &formattedData{Paragraphs: paragraphs}

		var buf bytes.Buffer
		if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
			t.Logf("writeFormattedDocumentXML failed: %v", err)
			return false
		}
		output := buf.String()

		// Verify exactly 1 <w:tbl>
		tblCount := strings.Count(output, `<w:tbl>`)
		if tblCount != 1 {
			t.Logf("expected 1 <w:tbl>, got %d", tblCount)
			return false
		}

		// Verify correct number of <w:tr>
		trCount := strings.Count(output, `<w:tr>`)
		if trCount != numRows {
			t.Logf("expected %d <w:tr>, got %d", numRows, trCount)
			return false
		}

		// Verify correct number of <w:tc>
		tcCount := strings.Count(output, `<w:tc>`)
		if tcCount != totalCells {
			t.Logf("expected %d <w:tc>, got %d", totalCells, tcCount)
			return false
		}

		// Verify <w:tblPr> is present
		if !strings.Contains(output, `<w:tblPr>`) {
			t.Logf("missing <w:tblPr>")
			return false
		}

		// Verify <w:tblBorders> is present
		if !strings.Contains(output, `<w:tblBorders>`) {
			t.Logf("missing <w:tblBorders>")
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: table OOXML output: %v", err)
	}
}

// TestWriteFormattedDocumentXML_Table verifies table XML structure is correct.
// _需求: 13.1, 13.2, 13.3, 13.4_
func TestWriteFormattedDocumentXML_Table(t *testing.T) {
	// 2 rows: row 1 has 2 cells, row 2 has 3 cells
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{InTable: true, IsTableCellEnd: true, Runs: []formattedRun{{Text: "A1", Props: doc.CharacterFormatting{}}}},
			{InTable: true, IsTableCellEnd: true, Runs: []formattedRun{{Text: "B1", Props: doc.CharacterFormatting{}}}},
			{InTable: true, TableRowEnd: true},
			{InTable: true, IsTableCellEnd: true, Runs: []formattedRun{{Text: "A2", Props: doc.CharacterFormatting{}}}},
			{InTable: true, IsTableCellEnd: true, Runs: []formattedRun{{Text: "B2", Props: doc.CharacterFormatting{}}}},
			{InTable: true, IsTableCellEnd: true, Runs: []formattedRun{{Text: "C2", Props: doc.CharacterFormatting{}}}},
			{InTable: true, TableRowEnd: true},
		},
	}

	var buf bytes.Buffer
	if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
		t.Fatalf("writeFormattedDocumentXML failed: %v", err)
	}
	output := buf.String()

	// Verify table structure
	if strings.Count(output, `<w:tbl>`) != 1 {
		t.Error("expected exactly 1 <w:tbl>")
	}
	if strings.Count(output, `</w:tbl>`) != 1 {
		t.Error("expected exactly 1 </w:tbl>")
	}
	if strings.Count(output, `<w:tr>`) != 2 {
		t.Errorf("expected 2 <w:tr>, got %d", strings.Count(output, `<w:tr>`))
	}
	if strings.Count(output, `<w:tc>`) != 5 {
		t.Errorf("expected 5 <w:tc>, got %d", strings.Count(output, `<w:tc>`))
	}

	// Verify cell content is present
	for _, text := range []string{"A1", "B1", "A2", "B2", "C2"} {
		if !strings.Contains(output, text) {
			t.Errorf("missing cell content %q", text)
		}
	}

	// Verify <w:tblPr> and <w:tblBorders> are present
	if !strings.Contains(output, `<w:tblPr>`) {
		t.Error("missing <w:tblPr>")
	}
	if !strings.Contains(output, `<w:tblBorders>`) {
		t.Error("missing <w:tblBorders>")
	}
}

// TestWriteFormattedDocumentXML_TableBorders verifies default borders are generated.
// _需求: 13.1, 13.2, 13.3, 13.4_
func TestWriteFormattedDocumentXML_TableBorders(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{InTable: true, IsTableCellEnd: true, Runs: []formattedRun{{Text: "cell", Props: doc.CharacterFormatting{}}}},
			{InTable: true, TableRowEnd: true},
		},
	}

	var buf bytes.Buffer
	if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
		t.Fatalf("writeFormattedDocumentXML failed: %v", err)
	}
	output := buf.String()

	// Verify all border sides are present
	borders := []string{"top", "left", "bottom", "right", "insideH", "insideV"}
	for _, side := range borders {
		expected := fmt.Sprintf(`<w:%s w:val="single" w:sz="4" w:space="0" w:color="auto"/>`, side)
		if !strings.Contains(output, expected) {
			t.Errorf("missing border: %s", expected)
		}
	}
}

// **Feature: doc-format-preservation, Property 11: 列表 OOXML 输出**
// For any set of list paragraphs (with random list types and levels), the generated
// DOCX should contain correct <w:numPr> elements, numbering.xml with decimal for
// ordered and bullet for unordered, and correct references in document.xml.rels
// and [Content_Types].xml.
// **Validates: Requirements 14.1, 14.2, 14.3, 14.4, 14.5, 14.6**
func TestProperty_ListOOXMLOutput(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		// Generate 1-10 list paragraphs with random types and levels
		numParas := 1 + rng.Intn(10)
		var paragraphs []formattedParagraph
		usedTypes := make(map[uint8]bool)

		for i := 0; i < numParas; i++ {
			listType := uint8(rng.Intn(2)) // 0=unordered, 1=ordered
			listLevel := uint8(rng.Intn(9)) // 0-8
			usedTypes[listType] = true
			paragraphs = append(paragraphs, formattedParagraph{
				IsListItem: true,
				ListType:   listType,
				ListLevel:  listLevel,
				Runs: []formattedRun{
					{Text: fmt.Sprintf("item%d", i), Props: doc.CharacterFormatting{}},
				},
			})
		}

		fd := &formattedData{Paragraphs: paragraphs}

		// Write full DOCX zip
		var zipBuf bytes.Buffer
		if err := writeDocxFormatted(&zipBuf, fd, nil); err != nil {
			t.Logf("writeDocxFormatted failed: %v", err)
			return false
		}

		// Read the zip
		zr, err := zip.NewReader(bytes.NewReader(zipBuf.Bytes()), int64(zipBuf.Len()))
		if err != nil {
			t.Logf("failed to read zip: %v", err)
			return false
		}

		// Read document.xml
		docXML, err := docReadZipFile(zr, "word/document.xml")
		if err != nil {
			t.Logf("failed to read document.xml: %v", err)
			return false
		}

		// Verify each list paragraph has <w:numPr> with correct <w:ilvl>
		listInfo := buildListNumInfo(fd)
		for i, p := range paragraphs {
			expectedIlvl := fmt.Sprintf(`<w:ilvl w:val="%d"/>`, p.ListLevel)
			if !strings.Contains(docXML, expectedIlvl) {
				t.Logf("missing ilvl for level %d", p.ListLevel)
				return false
			}
			if numId, ok := listInfo.paraNumId[i]; ok {
				expectedNumId := fmt.Sprintf(`<w:numId w:val="%d"/>`, numId)
				if !strings.Contains(docXML, expectedNumId) {
					t.Logf("missing numId %d for paragraph %d", numId, i)
					return false
				}
			} else {
				t.Logf("no numId assigned for paragraph %d", i)
				return false
			}
		}

		// Verify <w:numPr> count matches number of list paragraphs
		numPrCount := strings.Count(docXML, `<w:numPr>`)
		if numPrCount != numParas {
			t.Logf("expected %d <w:numPr>, got %d", numParas, numPrCount)
			return false
		}

		// Read numbering.xml
		numberingXML, err := docReadZipFile(zr, "word/numbering.xml")
		if err != nil {
			t.Logf("failed to read numbering.xml: %v", err)
			return false
		}

		// Verify numbering formats
		if usedTypes[0] {
			if !strings.Contains(numberingXML, `<w:numFmt w:val="bullet"/>`) {
				t.Logf("missing bullet format in numbering.xml")
				return false
			}
		}
		if usedTypes[1] {
			if !strings.Contains(numberingXML, `<w:numFmt w:val="decimal"/>`) {
				t.Logf("missing decimal format in numbering.xml")
				return false
			}
		}

		// Verify abstractNum count matches total numbering groups
		abstractNumCount := strings.Count(numberingXML, `<w:abstractNum `)
		if abstractNumCount != listInfo.totalNums {
			t.Logf("expected %d abstractNum, got %d", listInfo.totalNums, abstractNumCount)
			return false
		}

		// Verify num count matches total numbering groups
		numCount := strings.Count(numberingXML, `<w:num `)
		if numCount != listInfo.totalNums {
			t.Logf("expected %d num, got %d", listInfo.totalNums, numCount)
			return false
		}

		// Read document.xml.rels
		relsXML, err := docReadZipFile(zr, "word/_rels/document.xml.rels")
		if err != nil {
			t.Logf("failed to read document.xml.rels: %v", err)
			return false
		}

		// Verify numbering relationship
		if !strings.Contains(relsXML, `numbering.xml`) {
			t.Logf("missing numbering.xml relationship in document.xml.rels")
			return false
		}
		if !strings.Contains(relsXML, `relationships/numbering`) {
			t.Logf("missing numbering relationship type in document.xml.rels")
			return false
		}

		// Read [Content_Types].xml
		contentTypesXML, err := docReadZipFile(zr, "[Content_Types].xml")
		if err != nil {
			t.Logf("failed to read [Content_Types].xml: %v", err)
			return false
		}

		// Verify numbering content type
		if !strings.Contains(contentTypesXML, `numbering.xml`) {
			t.Logf("missing numbering.xml in [Content_Types].xml")
			return false
		}
		if !strings.Contains(contentTypesXML, `numbering+xml`) {
			t.Logf("missing numbering content type in [Content_Types].xml")
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: list OOXML output: %v", err)
	}
}

// TestWriteNumberingXML_Ordered verifies ordered lists use decimal format.
// _需求: 14.2, 14.3_
func TestWriteNumberingXML_Ordered(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{IsListItem: true, ListType: 1, ListLevel: 0, Runs: []formattedRun{{Text: "ordered item"}}},
		},
	}

	var buf bytes.Buffer
	if err := writeNumberingXML(&buf, fd); err != nil {
		t.Fatalf("writeNumberingXML failed: %v", err)
	}
	output := buf.String()

	if !strings.Contains(output, `<w:numFmt w:val="decimal"/>`) {
		t.Error("expected decimal format for ordered list")
	}
	if strings.Contains(output, `<w:numFmt w:val="bullet"/>`) {
		t.Error("unexpected bullet format for ordered-only list")
	}
	if !strings.Contains(output, `<w:abstractNum`) {
		t.Error("missing <w:abstractNum>")
	}
	if !strings.Contains(output, `<w:num `) {
		t.Error("missing <w:num>")
	}
}

// TestWriteNumberingXML_Unordered verifies unordered lists use bullet format.
// _需求: 14.2, 14.4_
func TestWriteNumberingXML_Unordered(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{IsListItem: true, ListType: 0, ListLevel: 0, Runs: []formattedRun{{Text: "bullet item"}}},
		},
	}

	var buf bytes.Buffer
	if err := writeNumberingXML(&buf, fd); err != nil {
		t.Fatalf("writeNumberingXML failed: %v", err)
	}
	output := buf.String()

	if !strings.Contains(output, `<w:numFmt w:val="bullet"/>`) {
		t.Error("expected bullet format for unordered list")
	}
	if strings.Contains(output, `<w:numFmt w:val="decimal"/>`) {
		t.Error("unexpected decimal format for unordered-only list")
	}
}

// TestWriteFormattedDocumentXML_ListParagraph verifies list paragraph <w:numPr> is correct.
// _需求: 14.1_
func TestWriteFormattedDocumentXML_ListParagraph(t *testing.T) {
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{
				IsListItem: true,
				ListType:   0, // unordered
				ListLevel:  2,
				Runs:       []formattedRun{{Text: "bullet item", Props: doc.CharacterFormatting{}}},
			},
			{
				IsListItem: true,
				ListType:   1, // ordered
				ListLevel:  0,
				Runs:       []formattedRun{{Text: "numbered item", Props: doc.CharacterFormatting{}}},
			},
			{
				IsListItem: false,
				Runs:       []formattedRun{{Text: "normal paragraph", Props: doc.CharacterFormatting{}}},
			},
		},
	}

	var buf bytes.Buffer
	if err := writeFormattedDocumentXML(&buf, fd, nil, nil); err != nil {
		t.Fatalf("writeFormattedDocumentXML failed: %v", err)
	}
	output := buf.String()

	// Verify <w:numPr> count: 2 list items
	numPrCount := strings.Count(output, `<w:numPr>`)
	if numPrCount != 2 {
		t.Errorf("expected 2 <w:numPr>, got %d", numPrCount)
	}

	// Verify ilvl for level 2
	if !strings.Contains(output, `<w:ilvl w:val="2"/>`) {
		t.Error("missing <w:ilvl w:val=\"2\"/>")
	}

	// Verify ilvl for level 0
	if !strings.Contains(output, `<w:ilvl w:val="0"/>`) {
		t.Error("missing <w:ilvl w:val=\"0\"/>")
	}

	// Verify numId references exist
	if !strings.Contains(output, `<w:numId w:val="`) {
		t.Error("missing <w:numId> element")
	}

	// Normal paragraph should not have <w:numPr>
	// The third paragraph should not contribute to numPr count (already checked above)
}

// TestWriteDocxFormatted_NumberingRels verifies numbering.xml relationship and content type.
// _需求: 14.5, 14.6_
func TestWriteDocxFormatted_NumberingRels(t *testing.T) {
	// With list paragraphs
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			{IsListItem: true, ListType: 0, ListLevel: 0, Runs: []formattedRun{{Text: "item"}}},
		},
	}

	var zipBuf bytes.Buffer
	if err := writeDocxFormatted(&zipBuf, fd, nil); err != nil {
		t.Fatalf("writeDocxFormatted failed: %v", err)
	}

	zr, err := zip.NewReader(bytes.NewReader(zipBuf.Bytes()), int64(zipBuf.Len()))
	if err != nil {
		t.Fatalf("failed to read zip: %v", err)
	}

	// Verify numbering.xml exists in the zip
	numberingXML, err := docReadZipFile(zr, "word/numbering.xml")
	if err != nil {
		t.Fatalf("numbering.xml not found in zip: %v", err)
	}
	if numberingXML == "" {
		t.Error("numbering.xml is empty")
	}

	// Verify document.xml.rels contains numbering relationship
	relsXML, err := docReadZipFile(zr, "word/_rels/document.xml.rels")
	if err != nil {
		t.Fatalf("failed to read document.xml.rels: %v", err)
	}
	if !strings.Contains(relsXML, `Target="numbering.xml"`) {
		t.Error("missing numbering.xml target in document.xml.rels")
	}
	if !strings.Contains(relsXML, `relationships/numbering`) {
		t.Error("missing numbering relationship type")
	}

	// Verify [Content_Types].xml contains numbering override
	contentTypesXML, err := docReadZipFile(zr, "[Content_Types].xml")
	if err != nil {
		t.Fatalf("failed to read [Content_Types].xml: %v", err)
	}
	if !strings.Contains(contentTypesXML, `/word/numbering.xml`) {
		t.Error("missing numbering.xml in [Content_Types].xml")
	}
	if !strings.Contains(contentTypesXML, `numbering+xml`) {
		t.Error("missing numbering content type")
	}

	// Without list paragraphs - numbering.xml should NOT be included
	fdNoList := &formattedData{
		Paragraphs: []formattedParagraph{
			{Runs: []formattedRun{{Text: "normal"}}},
		},
	}

	var zipBuf2 bytes.Buffer
	if err := writeDocxFormatted(&zipBuf2, fdNoList, nil); err != nil {
		t.Fatalf("writeDocxFormatted failed: %v", err)
	}

	zr2, err := zip.NewReader(bytes.NewReader(zipBuf2.Bytes()), int64(zipBuf2.Len()))
	if err != nil {
		t.Fatalf("failed to read zip: %v", err)
	}

	// numbering.xml should not exist
	_, err = docReadZipFile(zr2, "word/numbering.xml")
	if err == nil {
		t.Error("numbering.xml should not exist when there are no list paragraphs")
	}

	// document.xml.rels should not contain numbering reference
	relsXML2, err := docReadZipFile(zr2, "word/_rels/document.xml.rels")
	if err != nil {
		t.Fatalf("failed to read document.xml.rels: %v", err)
	}
	if strings.Contains(relsXML2, `numbering`) {
		t.Error("document.xml.rels should not contain numbering reference when no lists")
	}

	// [Content_Types].xml should not contain numbering override
	contentTypesXML2, err := docReadZipFile(zr2, "[Content_Types].xml")
	if err != nil {
		t.Fatalf("failed to read [Content_Types].xml: %v", err)
	}
	if strings.Contains(contentTypesXML2, `numbering`) {
		t.Error("[Content_Types].xml should not contain numbering when no lists")
	}
}

// TestIntegration_FormattedDocxOutput verifies the complete formatted DOCX output path
// by constructing diverse formatting data, generating a DOCX zip via writeDocxFormatted,
// and verifying that all expected OOXML elements are present.
// _需求: 16.3, 16.4, 16.5, 17.1, 17.2, 17.3_
func TestIntegration_FormattedDocxOutput(t *testing.T) {
	// Construct formattedData with diverse formatting:
	// - bold text run
	// - heading paragraph
	// - table paragraphs (cell + row end)
	// - list paragraph (ordered)
	// - page break paragraph
	fd := &formattedData{
		Paragraphs: []formattedParagraph{
			// 1. Bold text paragraph
			{
				Runs: []formattedRun{
					{Text: "Bold text", Props: doc.CharacterFormatting{Bold: true, FontName: "Arial", FontSize: 24}},
				},
				Props: doc.ParagraphFormatting{Alignment: 1}, // center
			},
			// 2. Heading paragraph (level 1)
			{
				HeadingLevel: 1,
				Runs:         []formattedRun{{Text: "Heading One"}},
			},
			// 3. Table: cell 1 in row 1
			{
				InTable: true,
				Runs:    []formattedRun{{Text: "Cell A1"}},
			},
			// 4. Table: cell 2 in row 1
			{
				InTable: true,
				Runs:    []formattedRun{{Text: "Cell B1"}},
			},
			// 5. Table: row end marker for row 1
			{
				InTable:     true,
				TableRowEnd: true,
				Runs:        []formattedRun{{Text: ""}},
			},
			// 6. Table: cell 1 in row 2
			{
				InTable: true,
				Runs:    []formattedRun{{Text: "Cell A2"}},
			},
			// 7. Table: row end marker for row 2
			{
				InTable:     true,
				TableRowEnd: true,
				Runs:        []formattedRun{{Text: ""}},
			},
			// 8. List paragraph (ordered, level 0)
			{
				IsListItem: true,
				ListType:   1, // ordered
				ListLevel:  0,
				Runs:       []formattedRun{{Text: "First item"}},
			},
			// 9. List paragraph (unordered, level 1)
			{
				IsListItem: true,
				ListType:   0, // unordered
				ListLevel:  1,
				Runs:       []formattedRun{{Text: "Bullet item"}},
			},
			// 10. Page break paragraph
			{
				PageBreakBefore: true,
				Runs:            []formattedRun{{Text: "After page break"}},
			},
		},
	}

	// Generate DOCX zip
	var zipBuf bytes.Buffer
	if err := writeDocxFormatted(&zipBuf, fd, nil); err != nil {
		t.Fatalf("writeDocxFormatted failed: %v", err)
	}

	// Parse the zip
	zr, err := zip.NewReader(bytes.NewReader(zipBuf.Bytes()), int64(zipBuf.Len()))
	if err != nil {
		t.Fatalf("failed to read zip: %v", err)
	}

	// --- Verify document.xml ---
	docXML, err := docReadZipFile(zr, "word/document.xml")
	if err != nil {
		t.Fatalf("failed to read document.xml: %v", err)
	}

	// Bold: <w:b/>
	if !strings.Contains(docXML, "<w:b/>") {
		t.Error("document.xml missing <w:b/> for bold text")
	}

	// Heading style: <w:pStyle w:val="Heading1"/>
	if !strings.Contains(docXML, `<w:pStyle w:val="Heading1"/>`) {
		t.Error("document.xml missing <w:pStyle w:val=\"Heading1\"/> for heading")
	}

	// Table elements: <w:tbl>, <w:tr>, <w:tc>
	if !strings.Contains(docXML, "<w:tbl>") {
		t.Error("document.xml missing <w:tbl> for table")
	}
	if !strings.Contains(docXML, "<w:tr>") {
		t.Error("document.xml missing <w:tr> for table row")
	}
	if !strings.Contains(docXML, "<w:tc>") {
		t.Error("document.xml missing <w:tc> for table cell")
	}

	// List: <w:numPr>
	if !strings.Contains(docXML, "<w:numPr>") {
		t.Error("document.xml missing <w:numPr> for list paragraph")
	}

	// Page break: <w:pageBreakBefore/>
	if !strings.Contains(docXML, "<w:pageBreakBefore/>") {
		t.Error("document.xml missing <w:pageBreakBefore/> for page break")
	}

	// --- Verify numbering.xml exists and has definitions ---
	numberingXML, err := docReadZipFile(zr, "word/numbering.xml")
	if err != nil {
		t.Fatalf("numbering.xml not found in zip: %v", err)
	}
	if !strings.Contains(numberingXML, "<w:abstractNum") {
		t.Error("numbering.xml missing <w:abstractNum> definitions")
	}
	if !strings.Contains(numberingXML, "<w:num ") {
		t.Error("numbering.xml missing <w:num> definitions")
	}

	// --- Verify styles.xml contains heading style definitions ---
	stylesXML, err := docReadZipFile(zr, "word/styles.xml")
	if err != nil {
		t.Fatalf("failed to read styles.xml: %v", err)
	}
	if !strings.Contains(stylesXML, `w:styleId="Heading1"`) {
		t.Error("styles.xml missing Heading1 style definition")
	}

	// --- Verify document.xml.rels contains numbering relationship ---
	relsXML, err := docReadZipFile(zr, "word/_rels/document.xml.rels")
	if err != nil {
		t.Fatalf("failed to read document.xml.rels: %v", err)
	}
	if !strings.Contains(relsXML, `Target="numbering.xml"`) {
		t.Error("document.xml.rels missing numbering.xml relationship")
	}

	// --- Verify [Content_Types].xml contains numbering content type ---
	contentTypesXML, err := docReadZipFile(zr, "[Content_Types].xml")
	if err != nil {
		t.Fatalf("failed to read [Content_Types].xml: %v", err)
	}
	if !strings.Contains(contentTypesXML, `/word/numbering.xml`) {
		t.Error("[Content_Types].xml missing numbering.xml part name")
	}
	if !strings.Contains(contentTypesXML, `numbering+xml`) {
		t.Error("[Content_Types].xml missing numbering content type")
	}

	// --- Verify ConvertReader and ConvertFile signatures compile correctly ---
	// These are compile-time checks: if the signatures change, this test won't compile.
	var _ func(io.ReadSeeker, io.Writer) error = ConvertReader
	var _ func(string, string) error = ConvertFile
}
