package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	// Parse the DOC
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening doc: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Extract all text from DOC paragraphs
	var docTexts []string
	for _, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			// Skip image placeholders
			cleaned := strings.ReplaceAll(r.Text, "\x01", "")
			cleaned = strings.ReplaceAll(cleaned, "\x08", "")
			cleaned = strings.ReplaceAll(cleaned, "\t", "")
			text += cleaned
		}
		text = strings.TrimSpace(text)
		if text != "" {
			docTexts = append(docTexts, text)
		}
	}

	// Read the DOCX
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening docx: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	var docxContent string
	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			docxContent = string(data)
		}
	}

	// Check each DOC text is present in DOCX
	fmt.Println("=== Text presence check ===")
	missing := 0
	for i, text := range docTexts {
		// Truncate for display
		display := text
		if len(display) > 60 {
			display = display[:60] + "..."
		}
		if !strings.Contains(docxContent, text) {
			// Try shorter prefix
			prefix := text
			if len(prefix) > 20 {
				prefix = prefix[:20]
			}
			if strings.Contains(docxContent, prefix) {
				// Partial match - might be split across runs
				continue
			}
			fmt.Printf("  MISSING[%d]: %q\n", i, display)
			missing++
		}
	}
	fmt.Printf("\nTotal doc texts: %d, Missing in docx: %d\n", len(docTexts), missing)

	// Check table content specifically
	fmt.Println("\n=== Table content check ===")
	tableTexts := []string{}
	for _, p := range fc.Paragraphs {
		if p.InTable {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			text = strings.TrimSpace(text)
			if text != "" {
				tableTexts = append(tableTexts, text)
			}
		}
	}
	fmt.Printf("Table texts in doc: %d\n", len(tableTexts))
	tableMissing := 0
	for _, text := range tableTexts {
		if !strings.Contains(docxContent, text) {
			prefix := text
			if len(prefix) > 15 {
				prefix = prefix[:15]
			}
			if !strings.Contains(docxContent, prefix) {
				display := text
				if len(display) > 60 {
					display = display[:60] + "..."
				}
				fmt.Printf("  TABLE MISSING: %q\n", display)
				tableMissing++
			}
		}
	}
	fmt.Printf("Table texts missing: %d\n", tableMissing)

	// Check image references
	fmt.Println("\n=== Image check ===")
	fmt.Printf("DOC images: %d\n", len(d.GetImages()))
	
	// Count inline and drawn images in doc
	inlineCount := 0
	drawnCount := 0
	for _, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				inlineCount++
			}
		}
		if len(p.DrawnImages) > 0 {
			drawnCount += len(p.DrawnImages)
		}
	}
	fmt.Printf("DOC inline image refs: %d, drawn image refs: %d\n", inlineCount, drawnCount)
	
	// Count images in docx
	docxInline := strings.Count(docxContent, "<wp:inline")
	docxAnchor := strings.Count(docxContent, "<wp:anchor")
	fmt.Printf("DOCX inline images: %d, anchor images: %d\n", docxInline, docxAnchor)
}
