package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: inspect_docx <file.docx>")
		os.Exit(1)
	}

	r, err := zip.OpenReader(os.Args[1])
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	fmt.Println("=== DOCX CONTENTS ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Read document.xml
	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			analyzeDocumentXML(data)
		}
		if f.Name == "word/styles.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== STYLES.XML (%d bytes) ===\n", len(data))
			// Count style definitions
			styleCount := strings.Count(string(data), "<w:style ")
			fmt.Printf("  Style definitions: %d\n", styleCount)
		}
	}
}

func analyzeDocumentXML(data []byte) {
	content := string(data)

	// Count elements
	paraCount := strings.Count(content, "<w:p>") + strings.Count(content, "<w:p ")
	runCount := strings.Count(content, "<w:r>") + strings.Count(content, "<w:r ")
	tableCount := strings.Count(content, "<w:tbl>")
	headingCount := 0
	for i := 1; i <= 9; i++ {
		headingCount += strings.Count(content, fmt.Sprintf(`w:val="Heading%d"`, i))
	}
	imgCount := strings.Count(content, "<w:drawing>")
	pageBreakCount := strings.Count(content, `w:type="page"`)
	listCount := strings.Count(content, "<w:numPr>")
	boldCount := strings.Count(content, "<w:b/>")
	fontCount := strings.Count(content, "<w:rFonts ")
	spacingCount := strings.Count(content, "<w:spacing ")
	alignCount := strings.Count(content, "<w:jc ")
	indentCount := strings.Count(content, "<w:ind ")
	szCount := strings.Count(content, "<w:sz ")
	tocFieldCount := strings.Count(content, "TOC \\o")
	tocStyleCount := 0
	for i := 1; i <= 3; i++ {
		tocStyleCount += strings.Count(content, fmt.Sprintf(`w:val="TOC%d"`, i))
	}
	sectionCount := strings.Count(content, "<w:sectPr")
	sectionBreakCount := strings.Count(content, `<w:type w:val="nextPage"`) +
		strings.Count(content, `<w:type w:val="continuous"`) +
		strings.Count(content, `<w:type w:val="evenPage"`) +
		strings.Count(content, `<w:type w:val="oddPage"`)
	headerRefCount := strings.Count(content, "<w:headerReference")
	footerRefCount := strings.Count(content, "<w:footerReference")

	fmt.Println("\n=== DOCUMENT.XML ANALYSIS ===")
	fmt.Printf("  Paragraphs: %d\n", paraCount)
	fmt.Printf("  Runs: %d\n", runCount)
	fmt.Printf("  Tables: %d\n", tableCount)
	fmt.Printf("  Headings: %d\n", headingCount)
	fmt.Printf("  Images: %d\n", imgCount)
	fmt.Printf("  Page breaks: %d\n", pageBreakCount)
	fmt.Printf("  List items: %d\n", listCount)
	fmt.Printf("  Bold runs: %d\n", boldCount)
	fmt.Printf("  Font specs: %d\n", fontCount)
	fmt.Printf("  Font sizes: %d\n", szCount)
	fmt.Printf("  Spacing rules: %d\n", spacingCount)
	fmt.Printf("  Alignments: %d\n", alignCount)
	fmt.Printf("  Indents: %d\n", indentCount)
	fmt.Printf("  TOC fields: %d\n", tocFieldCount)
	fmt.Printf("  TOC styled paras: %d\n", tocStyleCount)
	fmt.Printf("  Section properties: %d\n", sectionCount)
	fmt.Printf("  Section breaks: %d\n", sectionBreakCount)
	fmt.Printf("  Header references: %d\n", headerRefCount)
	fmt.Printf("  Footer references: %d\n", footerRefCount)

	// Extract text content page by page (split on page breaks)
	fmt.Println("\n=== TEXT CONTENT (by page break) ===")
	extractTextFromXML(data)
}

type docParser struct {
	pages    []string
	current  strings.Builder
	pageNum  int
	inRun    bool
	inText   bool
	inPara   bool
}

func extractTextFromXML(data []byte) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	var textBuf strings.Builder
	var pages []string
	inText := false
	paraText := ""
	paraCount := 0

	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := tok.(type) {
		case xml.StartElement:
			localName := t.Name.Local
			if localName == "t" {
				inText = true
			}
			if localName == "p" {
				paraText = ""
			}
		case xml.EndElement:
			localName := t.Name.Local
			if localName == "t" {
				inText = false
			}
			if localName == "p" {
				paraCount++
				if paraText != "" {
					textBuf.WriteString(paraText)
					textBuf.WriteString("\n")
				}
			}
		case xml.CharData:
			if inText {
				s := string(t)
				paraText += s
				// Check for page break marker
				if strings.Contains(s, "\f") {
					pages = append(pages, textBuf.String())
					textBuf.Reset()
				}
			}
		}
	}

	// Add remaining text as last page
	if textBuf.Len() > 0 {
		pages = append(pages, textBuf.String())
	}

	// If no page breaks found, just show all text
	if len(pages) == 0 {
		fmt.Printf("  (No page breaks detected, %d paragraphs total)\n", paraCount)
		return
	}

	for i, page := range pages {
		lines := strings.Split(strings.TrimSpace(page), "\n")
		fmt.Printf("\n--- Page %d (%d lines) ---\n", i+1, len(lines))
		for j, line := range lines {
			if j >= 5 {
				fmt.Printf("  ... (%d more lines)\n", len(lines)-5)
				break
			}
			if len(line) > 80 {
				line = line[:80] + "..."
			}
			fmt.Printf("  %s\n", line)
		}
	}
}
