package pptconv

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
)

func TestConvertReader_InvalidInput(t *testing.T) {
	// Provide invalid (non-PPT) data
	input := bytes.NewReader([]byte("this is not a valid PPT file"))
	var output bytes.Buffer

	err := ConvertReader(input, &output)
	if err == nil {
		t.Fatal("expected error for invalid input, got nil")
	}
	if !strings.Contains(err.Error(), "pptconv") {
		t.Errorf("error message should contain 'pptconv' prefix, got: %s", err.Error())
	}
}

func TestConvertFile_NonexistentInput(t *testing.T) {
	err := ConvertFile("/nonexistent/path/to/file.ppt", "/tmp/output.pptx")
	if err == nil {
		t.Fatal("expected error for nonexistent input file, got nil")
	}
	if !strings.Contains(err.Error(), "pptconv") {
		t.Errorf("error message should contain 'pptconv' prefix, got: %s", err.Error())
	}
}

func TestConvertFile_InvalidOutputPath(t *testing.T) {
	// Create a temp file with invalid PPT content to test the output path error
	// Since the input will fail to parse first, we test that error path
	err := ConvertFile("/nonexistent/input.ppt", "/nonexistent/dir/output.pptx")
	if err == nil {
		t.Fatal("expected error, got nil")
	}
	if !strings.Contains(err.Error(), "pptconv") {
		t.Errorf("error message should contain 'pptconv' prefix, got: %s", err.Error())
	}
}

// Feature: legacy-to-ooxml-conversion, Property 1: PPT 幻灯片文本保留
// **Validates: Requirements 1.3**
//
// For any set of slides with arbitrary text, after mapping through writePptx,
// the output PPTX should contain all slide texts in order.
func TestProperty_SlideTextPreservation(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(slideCount uint8, seed int64) bool {
		// Limit slide count to a reasonable range (0-10)
		numSlides := int(slideCount) % 11
		rng := rand.New(rand.NewSource(seed))

		// Generate random slides with random texts
		slides := make([]slideData, numSlides)
		for i := 0; i < numSlides; i++ {
			numTexts := rng.Intn(5) // 0-4 texts per slide
			texts := make([]string, numTexts)
			for j := 0; j < numTexts; j++ {
				texts[j] = randomString(rng, 1+rng.Intn(50))
			}
			slides[i] = slideData{Texts: texts}
		}

		// Write PPTX to buffer
		var buf bytes.Buffer
		if err := writePptx(&buf, slides, nil); err != nil {
			t.Logf("writePptx failed: %v", err)
			return false
		}

		// Read back the zip and extract slide texts
		reader := bytes.NewReader(buf.Bytes())
		zr, err := zip.NewReader(reader, int64(buf.Len()))
		if err != nil {
			t.Logf("failed to open zip: %v", err)
			return false
		}

		// For each slide, verify texts are present in the slide XML
		for i, slide := range slides {
			slideFile := fmt.Sprintf("ppt/slides/slide%d.xml", i+1)
			content, err := readZipFile(zr, slideFile)
			if err != nil {
				t.Logf("failed to read %s: %v", slideFile, err)
				return false
			}

			for _, text := range slide.Texts {
				escaped := escapeXML(text)
				if !strings.Contains(content, escaped) {
					t.Logf("slide %d missing text %q (escaped: %q) in XML", i+1, text, escaped)
					return false
				}
			}
		}

		// Verify slide count
		if numSlides > 0 {
			slideCount := countSlideFiles(zr)
			if slideCount != numSlides {
				t.Logf("expected %d slides, got %d", numSlides, slideCount)
				return false
			}
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: PPT slide text preservation: %v", err)
	}
}

// randomString generates a random alphanumeric string of the given length.
func randomString(rng *rand.Rand, length int) string {
	const charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 "
	b := make([]byte, length)
	for i := range b {
		b[i] = charset[rng.Intn(len(charset))]
	}
	return string(b)
}

// readZipFile reads the content of a file inside a zip archive.
func readZipFile(zr *zip.Reader, name string) (string, error) {
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

// countSlideFiles counts the number of slide XML files in the zip.
func countSlideFiles(zr *zip.Reader) int {
	count := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			count++
		}
	}
	return count
}

// escapeXML escapes a string the same way xml.Escape does.
func escapeXML(s string) string {
	var buf bytes.Buffer
	xml.Escape(&buf, []byte(s))
	return buf.String()
}

// Feature: legacy-to-ooxml-conversion, Property 2: PPT 图片保留
// **Validates: Requirements 1.4**
//
// For any set of images with arbitrary format and data, after mapping through writePptx,
// the output PPTX should contain all images with matching data.
func TestProperty_ImagePreservation(t *testing.T) {
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

		// Write PPTX to buffer (with one empty slide so it's valid)
		slides := []slideData{{Texts: []string{}}}
		var buf bytes.Buffer
		if err := writePptx(&buf, slides, images); err != nil {
			t.Logf("writePptx failed: %v", err)
			return false
		}

		// Read back the zip and verify images
		reader := bytes.NewReader(buf.Bytes())
		zr, err := zip.NewReader(reader, int64(buf.Len()))
		if err != nil {
			t.Logf("failed to open zip: %v", err)
			return false
		}

		// Count image files in ppt/media/
		imageFiles := 0
		for _, f := range zr.File {
			if strings.HasPrefix(f.Name, "ppt/media/") {
				imageFiles++
			}
		}

		if imageFiles != numImages {
			t.Logf("expected %d images, got %d", numImages, imageFiles)
			return false
		}

		// Verify each image's data matches
		for i, img := range images {
			ext := (&common.Image{Format: img.Format}).Extension()
			if ext == "" {
				ext = ".bin"
			}
			filename := fmt.Sprintf("ppt/media/image%d%s", i+1, ext)
			content, err := readZipFileBytes(zr, filename)
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
		t.Errorf("Property failed: PPT image preservation: %v", err)
	}
}

// readZipFileBytes reads the raw bytes of a file inside a zip archive.
func readZipFileBytes(zr *zip.Reader, name string) ([]byte, error) {
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

// Feature: ppt-to-pptx-format-conversion, Property 5: Character format PresentationML output
// For any SlideTextRun with random font name, font size, bold, italic, underline, color,
// the generated <a:rPr> XML should contain the corresponding PresentationML elements.
func TestProperty_CharFormatOutput(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		run := formattedSlideRun{
			Text:      "test",
			FontName:  randomString(rng, 1+rng.Intn(20)),
			FontSize:  uint16(100 + rng.Intn(10000)),
			Bold:      rng.Intn(2) == 1,
			Italic:    rng.Intn(2) == 1,
			Underline: rng.Intn(2) == 1,
			Color:     fmt.Sprintf("%02X%02X%02X", rng.Intn(256), rng.Intn(256), rng.Intn(256)),
		}

		var buf bytes.Buffer
		writeSlideRunProperties(&buf, run)
		xml := buf.String()

		// Check font size
		if !strings.Contains(xml, fmt.Sprintf(`sz="%d"`, run.FontSize)) {
			t.Logf("missing sz=%d in %s", run.FontSize, xml)
			return false
		}

		// Check bold
		if run.Bold && !strings.Contains(xml, `b="1"`) {
			t.Logf("missing b=1 for bold")
			return false
		}

		// Check italic
		if run.Italic && !strings.Contains(xml, `i="1"`) {
			t.Logf("missing i=1 for italic")
			return false
		}

		// Check underline
		if run.Underline && !strings.Contains(xml, `u="sng"`) {
			t.Logf("missing u=sng for underline")
			return false
		}

		// Check color
		if !strings.Contains(xml, fmt.Sprintf(`<a:srgbClr val="%s"/>`, run.Color)) {
			t.Logf("missing color %s", run.Color)
			return false
		}

		// Check font name
		escaped := xmlEscapeAttr(run.FontName)
		if !strings.Contains(xml, fmt.Sprintf(`<a:latin typeface="%s"/>`, escaped)) {
			t.Logf("missing latin typeface %s", run.FontName)
			return false
		}
		if !strings.Contains(xml, fmt.Sprintf(`<a:ea typeface="%s"/>`, escaped)) {
			t.Logf("missing ea typeface %s", run.FontName)
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: char format output: %v", err)
	}
}

// Feature: ppt-to-pptx-format-conversion, Property 6: Paragraph format PresentationML output
// For any SlideParagraph with random alignment, indent level, spacing,
// the generated <a:pPr> XML should contain the corresponding PresentationML elements.
func TestProperty_ParaFormatOutput(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		alignment := uint8(1 + rng.Intn(3)) // 1-3 (skip 0 which is default/left)
		indentLevel := uint8(1 + rng.Intn(4))
		// Negative values = absolute centipoints, positive = percentage
		spaceBefore := -int32(1 + rng.Intn(500))
		spaceAfter := -int32(1 + rng.Intn(500))
		lineSpacing := -int32(1 + rng.Intn(500))

		para := formattedSlideParagraph{
			Alignment:   alignment,
			IndentLevel: indentLevel,
			SpaceBefore: spaceBefore,
			SpaceAfter:  spaceAfter,
			LineSpacing: lineSpacing,
			Runs:        []formattedSlideRun{{Text: "test"}},
		}

		var buf bytes.Buffer
		writeTextBodyXML(&buf, []formattedSlideParagraph{para})
		xml := buf.String()

		// Check alignment
		algnMap := []string{"l", "ctr", "r", "just"}
		if int(alignment) < len(algnMap) {
			expected := fmt.Sprintf(`algn="%s"`, algnMap[alignment])
			if !strings.Contains(xml, expected) {
				t.Logf("missing %s in %s", expected, xml)
				return false
			}
		}

		// Check indent level
		if !strings.Contains(xml, fmt.Sprintf(`lvl="%d"`, indentLevel)) {
			t.Logf("missing lvl=%d", indentLevel)
			return false
		}

		// Negative values → absolute centipoints → spcPts with positive value
		if !strings.Contains(xml, fmt.Sprintf(`<a:lnSpc><a:spcPts val="%d"/></a:lnSpc>`, -lineSpacing)) {
			t.Logf("missing lnSpc %d", -lineSpacing)
			return false
		}
		if !strings.Contains(xml, fmt.Sprintf(`<a:spcBef><a:spcPts val="%d"/></a:spcBef>`, -spaceBefore)) {
			t.Logf("missing spcBef %d", -spaceBefore)
			return false
		}
		if !strings.Contains(xml, fmt.Sprintf(`<a:spcAft><a:spcPts val="%d"/></a:spcAft>`, -spaceAfter)) {
			t.Logf("missing spcAft %d", -spaceAfter)
			return false
		}

		// Also test positive values (percentage mode)
		pctPara := formattedSlideParagraph{
			LineSpacing: 100, // 100% = single space
			SpaceBefore: 50,  // 50%
			Runs:        []formattedSlideRun{{Text: "test"}},
		}
		var buf2 bytes.Buffer
		writeTextBodyXML(&buf2, []formattedSlideParagraph{pctPara})
		xml2 := buf2.String()
		if !strings.Contains(xml2, `<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>`) {
			t.Logf("missing percentage lnSpc in %s", xml2)
			return false
		}
		if !strings.Contains(xml2, `<a:spcBef><a:spcPct val="50000"/></a:spcBef>`) {
			t.Logf("missing percentage spcBef in %s", xml2)
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: para format output: %v", err)
	}
}

// Feature: ppt-to-pptx-format-conversion, Property 7: Shape layout PresentationML output
// For any shape with random position and size, writeShapeXML should generate XML
// with <a:off> and <a:ext> elements matching the input EMU values.
func TestProperty_ShapeLayoutOutput(t *testing.T) {
	config := &quick.Config{MaxCount: 100}

	prop := func(seed int64) bool {
		rng := rand.New(rand.NewSource(seed))

		left := int32(rng.Intn(9144000))
		top := int32(rng.Intn(6858000))
		width := int32(1 + rng.Intn(5000000))
		height := int32(1 + rng.Intn(3000000))

		shape := formattedShape{
			ShapeType: 202,
			Left:      left,
			Top:       top,
			Width:     width,
			Height:    height,
			IsText:    true,
			Paragraphs: []formattedSlideParagraph{
				{Runs: []formattedSlideRun{{Text: "test"}}},
			},
		}

		var buf bytes.Buffer
		writeShapeXML(&buf, shape, 2, nil)
		xml := buf.String()

		expectedOff := fmt.Sprintf(`<a:off x="%d" y="%d"/>`, left, top)
		expectedExt := fmt.Sprintf(`<a:ext cx="%d" cy="%d"/>`, width, height)

		if !strings.Contains(xml, expectedOff) {
			t.Logf("missing offset: %s", expectedOff)
			return false
		}
		if !strings.Contains(xml, expectedExt) {
			t.Logf("missing extent: %s", expectedExt)
			return false
		}

		return true
	}

	if err := quick.Check(prop, config); err != nil {
		t.Errorf("Property failed: shape layout output: %v", err)
	}
}

// Test formatted slide generation produces valid XML structure
func TestWriteFormattedSlideXML(t *testing.T) {
	slide := formattedSlideData{
		Shapes: []formattedShape{
			{
				ShapeType: 202,
				Left:      100000,
				Top:       200000,
				Width:     5000000,
				Height:    1000000,
				IsText:    true,
				Paragraphs: []formattedSlideParagraph{
					{
						Alignment: 1, // center
						Runs: []formattedSlideRun{
							{
								Text:     "Hello",
								FontName: "Arial",
								FontSize: 2400,
								Bold:     true,
								Color:    "FF0000",
							},
						},
					},
				},
			},
		},
	}

	var buf bytes.Buffer
	err := writeFormattedSlideXML(&buf, slide, nil, nil)
	if err != nil {
		t.Fatalf("writeFormattedSlideXML failed: %v", err)
	}

	xml := buf.String()
	if !strings.Contains(xml, `<p:sld`) {
		t.Error("missing <p:sld> root element")
	}
	if !strings.Contains(xml, `b="1"`) {
		t.Error("missing bold attribute")
	}
	if !strings.Contains(xml, `sz="2400"`) {
		t.Error("missing font size")
	}
	if !strings.Contains(xml, `algn="ctr"`) {
		t.Error("missing center alignment")
	}
	if !strings.Contains(xml, `FF0000`) {
		t.Error("missing color")
	}
}
