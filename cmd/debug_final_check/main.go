package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer r.Close()

	fmt.Println("=== DOCX Structure ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			fmt.Printf("\n=== Document Stats ===\n")
			fmt.Printf("Total size: %d bytes\n", len(content))
			fmt.Printf("Paragraphs: %d\n", strings.Count(content, "<w:p>")+strings.Count(content, "<w:p "))
			fmt.Printf("Drawings: %d\n", strings.Count(content, "<w:drawing>"))
			fmt.Printf("Tables: %d\n", strings.Count(content, "<w:tbl>"))
			fmt.Printf("Page breaks: %d\n", strings.Count(content, `w:type="page"`))
			fmt.Printf("Headings: %d\n", strings.Count(content, "Heading"))
			fmt.Printf("List items: %d\n", strings.Count(content, "<w:numPr>"))
			fmt.Printf("Sections: %d\n", strings.Count(content, "<w:sectPr"))
			fmt.Printf("Header refs: %d\n", strings.Count(content, "headerReference"))
			fmt.Printf("Footer refs: %d\n", strings.Count(content, "footerReference"))
			fmt.Printf("Title page: %v\n", strings.Contains(content, "<w:titlePg/>"))
			
			// Check for common issues
			fmt.Printf("\n=== Quality Checks ===\n")
			
			// Check textbox
			fmt.Printf("Has textbox: %v\n", strings.Contains(content, "txBox"))
			
			// Check inline vs anchor images
			inlineCount := strings.Count(content, "<wp:inline")
			anchorCount := strings.Count(content, "<wp:anchor")
			fmt.Printf("Inline images: %d, Anchor images: %d\n", inlineCount, anchorCount)
			
			// Check for wrapTopAndBottom
			wrapTB := strings.Count(content, "wrapTopAndBottom")
			wrapNone := strings.Count(content, "wrapNone")
			fmt.Printf("Wrap top/bottom: %d, Wrap none: %d\n", wrapTB, wrapNone)
		}
		
		if f.Name == "word/settings.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== settings.xml ===\n%s\n", string(data))
		}
	}
}
