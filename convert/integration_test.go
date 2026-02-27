package convert_test

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
	"testing"

	"github.com/shakinm/xlsReader/convert/docconv"
	"github.com/shakinm/xlsReader/convert/pptconv"
)

func TestIntegrationPPTtoPPTX(t *testing.T) {
	inputPath := "../testfie/test.ppt"
	if _, err := os.Stat(inputPath); os.IsNotExist(err) {
		t.Fatalf("test fixture missing: %s", inputPath)
	}

	outPath := t.TempDir() + "/test.pptx"
	err := pptconv.ConvertFile(inputPath, outPath)
	if err != nil {
		t.Fatalf("ConvertFile failed: %v", err)
	}

	info, err := os.Stat(outPath)
	if err != nil {
		t.Fatalf("output file does not exist: %v", err)
	}
	if info.Size() == 0 {
		t.Fatal("output file is empty")
	}
	t.Logf("Output PPTX size: %d bytes", info.Size())

	// Open as zip and inspect structure
	zr, err := zip.OpenReader(outPath)
	if err != nil {
		t.Fatalf("failed to open output as zip: %v", err)
	}
	defer zr.Close()

	// List all files in the archive
	t.Log("=== PPTX archive contents ===")
	for _, f := range zr.File {
		t.Logf("  %s (%d bytes)", f.Name, f.UncompressedSize64)
	}

	// Verify required structure files exist
	requiredFiles := []string{
		"ppt/presentation.xml",
		"ppt/slides/slide1.xml",
		"[Content_Types].xml",
		"_rels/.rels",
	}
	for _, req := range requiredFiles {
		if findZipEntry(&zr.Reader, req) == nil {
			t.Errorf("missing required file in PPTX: %s", req)
		}
	}

	// Count slides
	slideCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") &&
			!strings.Contains(f.Name, "_rels") {
			slideCount++
		}
	}
	t.Logf("Slide count: %d", slideCount)

	// Extract and print text from each slide
	t.Log("=== Slide text content ===")
	for i := 1; i <= slideCount; i++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		content, err := readZipEntry(&zr.Reader, name)
		if err != nil {
			t.Errorf("failed to read %s: %v", name, err)
			continue
		}
		texts := extractPptxTexts(content)
		t.Logf("Slide %d texts: %v", i, texts)
	}
}

func TestIntegrationDOCtoDOCX(t *testing.T) {
	inputPath := "../testfie/test.doc"
	if _, err := os.Stat(inputPath); os.IsNotExist(err) {
		t.Fatalf("test fixture missing: %s", inputPath)
	}

	outPath := t.TempDir() + "/test.docx"
	err := docconv.ConvertFile(inputPath, outPath)
	if err != nil {
		t.Fatalf("ConvertFile failed: %v", err)
	}

	info, err := os.Stat(outPath)
	if err != nil {
		t.Fatalf("output file does not exist: %v", err)
	}
	if info.Size() == 0 {
		t.Fatal("output file is empty")
	}
	t.Logf("Output DOCX size: %d bytes", info.Size())

	// Open as zip and inspect structure
	zr, err := zip.OpenReader(outPath)
	if err != nil {
		t.Fatalf("failed to open output as zip: %v", err)
	}
	defer zr.Close()

	// List all files in the archive
	t.Log("=== DOCX archive contents ===")
	for _, f := range zr.File {
		t.Logf("  %s (%d bytes)", f.Name, f.UncompressedSize64)
	}

	// Verify required structure files exist
	requiredFiles := []string{
		"word/document.xml",
		"[Content_Types].xml",
		"_rels/.rels",
	}
	for _, req := range requiredFiles {
		if findZipEntry(&zr.Reader, req) == nil {
			t.Errorf("missing required file in DOCX: %s", req)
		}
	}

	// Extract and print text from document.xml
	content, err := readZipEntry(&zr.Reader, "word/document.xml")
	if err != nil {
		t.Fatalf("failed to read word/document.xml: %v", err)
	}

	text := extractDocxText(content)
	t.Log("=== Document text content ===")
	t.Logf("Text: %s", text)
}

// --- helpers ---

func findZipEntry(zr *zip.Reader, name string) *zip.File {
	for _, f := range zr.File {
		if f.Name == name {
			return f
		}
	}
	return nil
}

func readZipEntry(zr *zip.Reader, name string) ([]byte, error) {
	f := findZipEntry(zr, name)
	if f == nil {
		return nil, fmt.Errorf("file %s not found in zip", name)
	}
	rc, err := f.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

// extractPptxTexts pulls all <a:t> text nodes from a slide XML.
func extractPptxTexts(xmlData []byte) []string {
	type AText struct {
		Value string `xml:",chardata"`
	}
	decoder := xml.NewDecoder(strings.NewReader(string(xmlData)))
	var texts []string
	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "t" {
			var at AText
			if err := decoder.DecodeElement(&at, &se); err == nil && at.Value != "" {
				texts = append(texts, at.Value)
			}
		}
	}
	return texts
}

// extractDocxText pulls all <w:t> text nodes from document.xml.
func extractDocxText(xmlData []byte) string {
	type WText struct {
		Value string `xml:",chardata"`
	}
	decoder := xml.NewDecoder(strings.NewReader(string(xmlData)))
	var sb strings.Builder
	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "t" {
			var wt WText
			if err := decoder.DecodeElement(&wt, &se); err == nil {
				sb.WriteString(wt.Value)
			}
		}
	}
	return sb.String()
}
