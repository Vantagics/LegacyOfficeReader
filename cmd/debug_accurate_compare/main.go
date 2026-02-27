package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func extractAllText(content string) string {
	var result strings.Builder
	idx := 0
	for {
		tStart := strings.Index(content[idx:], "<w:t")
		if tStart < 0 {
			break
		}
		tStart += idx
		gt := strings.Index(content[tStart:], ">")
		if gt < 0 {
			break
		}
		textStart := tStart + gt + 1
		tEnd := strings.Index(content[textStart:], "</w:t>")
		if tEnd < 0 {
			break
		}
		text := content[textStart : textStart+tEnd]
		// Decode XML entities
		text = strings.ReplaceAll(text, "&amp;", "&")
		text = strings.ReplaceAll(text, "&lt;", "<")
		text = strings.ReplaceAll(text, "&gt;", ">")
		text = strings.ReplaceAll(text, "&#x9;", "\t")
		result.WriteString(text)
		idx = textStart + tEnd + 6
	}
	return result.String()
}

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

	// Extract all meaningful text from DOC (skip empty, skip image placeholders)
	var docParagraphTexts []string
	for _, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			cleaned := r.Text
			cleaned = strings.ReplaceAll(cleaned, "\x01", "")
			cleaned = strings.ReplaceAll(cleaned, "\x08", "")
			text += cleaned
		}
		text = strings.TrimSpace(text)
		if text != "" && len(text) > 2 {
			docParagraphTexts = append(docParagraphTexts, text)
		}
	}

	// Read the DOCX
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening docx: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	var docxAllText string
	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			docxAllText = extractAllText(string(data))
		}
	}

	// Check each DOC paragraph text is present in DOCX
	fmt.Println("=== Paragraph text comparison ===")
	missing := 0
	for _, text := range docParagraphTexts {
		// Remove tabs and spaces for comparison
		cleanDoc := strings.ReplaceAll(text, "\t", "")
		cleanDoc = strings.ReplaceAll(cleanDoc, " ", "")
		cleanDocx := strings.ReplaceAll(docxAllText, "\t", "")
		cleanDocx = strings.ReplaceAll(cleanDocx, " ", "")

		if !strings.Contains(cleanDocx, cleanDoc) {
			// Try with first 15 chars
			prefix := cleanDoc
			if len(prefix) > 15 {
				prefix = prefix[:15]
			}
			if !strings.Contains(cleanDocx, prefix) {
				display := text
				if len(display) > 80 {
					display = display[:80] + "..."
				}
				fmt.Printf("  MISSING: %q\n", display)
				missing++
			}
		}
	}
	fmt.Printf("\nTotal doc paragraphs with text: %d, Missing: %d\n", len(docParagraphTexts), missing)

	// Also check header/footer text
	fmt.Println("\n=== Header/Footer check ===")
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "word/header") || strings.HasPrefix(f.Name, "word/footer") {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			text := extractAllText(string(data))
			fmt.Printf("%s: %q\n", f.Name, text)
		}
	}
}
