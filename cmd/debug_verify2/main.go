package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Read document.xml
	var docXML string
	for _, f := range zr.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			docXML = string(data)
		}
	}

	// Count key elements
	fmt.Printf("=== DOCX Structure ===\n")
	fmt.Printf("Paragraphs (<w:p>): %d\n", strings.Count(docXML, "<w:p>"))
	fmt.Printf("Tables (<w:tbl>): %d\n", strings.Count(docXML, "<w:tbl>"))
	fmt.Printf("Table rows (<w:tr>): %d\n", strings.Count(docXML, "<w:tr>"))
	fmt.Printf("Table cells (<w:tc>): %d\n", strings.Count(docXML, "<w:tc>"))
	fmt.Printf("Images (<w:drawing>): %d\n", strings.Count(docXML, "<w:drawing>"))
	fmt.Printf("Page breaks (<w:br w:type=\"page\"/>): %d\n", strings.Count(docXML, `w:type="page"`))
	fmt.Printf("Section breaks (<w:sectPr>): %d\n", strings.Count(docXML, "<w:sectPr>"))
	fmt.Printf("Headings (Heading): %d\n", strings.Count(docXML, `w:val="Heading`))
	fmt.Printf("TOC entries (TOC): %d\n", strings.Count(docXML, `w:val="TOC`))
	fmt.Printf("List items (<w:numPr>): %d\n", strings.Count(docXML, "<w:numPr>"))
	fmt.Printf("Bold (<w:b/>): %d\n", strings.Count(docXML, "<w:b/>"))
	fmt.Printf("Center align: %d\n", strings.Count(docXML, `w:val="center"`))
	fmt.Printf("Right align: %d\n", strings.Count(docXML, `w:val="right"`))
	fmt.Printf("Both align: %d\n", strings.Count(docXML, `w:val="both"`))

	// Check image references
	fmt.Printf("\n=== Image References ===\n")
	for i := 1; i <= 8; i++ {
		ref := fmt.Sprintf("rImg%d", i)
		count := strings.Count(docXML, ref)
		fmt.Printf("  %s: %d references\n", ref, count)
	}

	// Show first 3000 chars of document.xml to check title page
	fmt.Printf("\n=== Document.xml first 3000 chars ===\n")
	if len(docXML) > 3000 {
		fmt.Println(docXML[:3000])
	} else {
		fmt.Println(docXML)
	}
}
