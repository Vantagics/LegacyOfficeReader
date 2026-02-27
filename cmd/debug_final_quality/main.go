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
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Read the DOCX
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	var docxContent string
	var settingsContent string
	var stylesContent string
	fileCount := 0
	for _, f := range r.File {
		fileCount++
		switch f.Name {
		case "word/document.xml":
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			docxContent = string(data)
		case "word/settings.xml":
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			settingsContent = string(data)
		case "word/styles.xml":
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			stylesContent = string(data)
		}
	}

	fmt.Println("=== DOCX Quality Report ===")
	fmt.Printf("ZIP files: %d\n", fileCount)
	issues := 0

	// 1. Check XML declaration
	if !strings.HasPrefix(docxContent, `<?xml version="1.0"`) {
		fmt.Println("ISSUE: Missing XML declaration")
		issues++
	}

	// 2. Check document structure
	if !strings.Contains(docxContent, "<w:body>") {
		fmt.Println("ISSUE: Missing w:body")
		issues++
	}
	if !strings.Contains(docxContent, "</w:body>") {
		fmt.Println("ISSUE: Missing closing w:body")
		issues++
	}

	// 3. Check section properties
	sectPrCount := strings.Count(docxContent, "<w:sectPr")
	fmt.Printf("Section properties: %d\n", sectPrCount)
	if sectPrCount < 1 {
		fmt.Println("ISSUE: No section properties")
		issues++
	}

	// 4. Check page size and margins
	if !strings.Contains(docxContent, `w:w="11906"`) || !strings.Contains(docxContent, `w:h="16838"`) {
		fmt.Println("ISSUE: Incorrect page size (should be A4)")
		issues++
	}

	// 5. Check titlePg
	if !strings.Contains(docxContent, "<w:titlePg/>") {
		fmt.Println("ISSUE: Missing titlePg for first page header")
		issues++
	}

	// 6. Check evenAndOddHeaders in settings
	if !strings.Contains(settingsContent, "evenAndOddHeaders") {
		fmt.Println("ISSUE: Missing evenAndOddHeaders in settings.xml")
		issues++
	}

	// 7. Check heading styles
	for i := 1; i <= 3; i++ {
		if !strings.Contains(stylesContent, fmt.Sprintf(`w:styleId="Heading%d"`, i)) {
			fmt.Printf("ISSUE: Missing Heading%d style\n", i)
			issues++
		}
	}

	// 8. Check TOC styles
	for i := 1; i <= 3; i++ {
		if !strings.Contains(stylesContent, fmt.Sprintf(`w:styleId="TOC%d"`, i)) {
			fmt.Printf("ISSUE: Missing TOC%d style\n", i)
			issues++
		}
	}

	// 9. Check header/footer styles
	if !strings.Contains(stylesContent, `w:styleId="Header"`) {
		fmt.Println("ISSUE: Missing Header style")
		issues++
	}
	if !strings.Contains(stylesContent, `w:styleId="Footer"`) {
		fmt.Println("ISSUE: Missing Footer style")
		issues++
	}

	// 10. Check content completeness
	// Count headings in doc
	docH1 := 0
	docH2 := 0
	docH3 := 0
	for _, p := range fc.Paragraphs {
		switch p.HeadingLevel {
		case 1:
			docH1++
		case 2:
			docH2++
		case 3:
			docH3++
		}
	}
	docxH1 := strings.Count(docxContent, `w:val="Heading1"`)
	docxH2 := strings.Count(docxContent, `w:val="Heading2"`)
	docxH3 := strings.Count(docxContent, `w:val="Heading3"`)
	fmt.Printf("Headings - DOC: H1=%d H2=%d H3=%d, DOCX: H1=%d H2=%d H3=%d\n",
		docH1, docH2, docH3, docxH1, docxH2, docxH3)
	if docH1 != docxH1 || docH2 != docxH2 || docH3 != docxH3 {
		fmt.Println("ISSUE: Heading count mismatch")
		issues++
	}

	// 11. Check images
	docImages := len(d.GetImages())
	docxImages := 0
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "word/media/") {
			docxImages++
		}
	}
	fmt.Printf("Images - DOC: %d, DOCX: %d\n", docImages, docxImages)
	if docxImages < docImages {
		fmt.Println("ISSUE: Missing images")
		issues++
	}

	// 12. Check table
	tblCount := strings.Count(docxContent, "<w:tbl>")
	trCount := strings.Count(docxContent, "<w:tr>")
	fmt.Printf("Tables: %d, Rows: %d\n", tblCount, trCount)

	// 13. Check TOC field
	if !strings.Contains(docxContent, "TOC \\o") {
		fmt.Println("ISSUE: Missing TOC field")
		issues++
	}

	// 14. Check page breaks
	pbCount := strings.Count(docxContent, `w:type="page"`)
	docPB := 0
	for _, p := range fc.Paragraphs {
		if p.HasPageBreak {
			docPB++
		}
	}
	fmt.Printf("Page breaks - DOC: %d, DOCX: %d\n", docPB, docPB)
	if pbCount != docPB {
		fmt.Printf("ISSUE: Page break count mismatch (docx has %d)\n", pbCount)
		issues++
	}

	// 15. Check list numbering
	numIdCount := strings.Count(docxContent, "<w:numId")
	docListItems := 0
	for _, p := range fc.Paragraphs {
		if p.IsListItem {
			docListItems++
		}
	}
	fmt.Printf("List items - DOC: %d, DOCX numId refs: %d\n", docListItems, numIdCount)

	// 16. Check header/footer references
	hdrRefCount := strings.Count(docxContent, "<w:headerReference")
	ftrRefCount := strings.Count(docxContent, "<w:footerReference")
	fmt.Printf("Header refs: %d, Footer refs: %d\n", hdrRefCount, ftrRefCount)
	fmt.Printf("Header entries: %d, Footer entries: %d\n", len(fc.HeaderEntries), len(fc.FooterEntries))

	// 17. Check text box
	if !strings.Contains(docxContent, "<wps:wsp>") {
		fmt.Println("ISSUE: Missing text box")
		issues++
	}

	// 18. Check Normal style default alignment
	if !strings.Contains(stylesContent, `<w:jc w:val="both"/>`) {
		fmt.Println("ISSUE: Normal style missing justify alignment")
		issues++
	}

	// 19. Check document defaults
	if !strings.Contains(stylesContent, `w:eastAsia="宋体"`) {
		fmt.Println("ISSUE: Missing 宋体 as default eastAsia font")
		issues++
	}

	fmt.Printf("\n=== Total issues: %d ===\n", issues)
	if issues == 0 {
		fmt.Println("All checks passed!")
	}
}
