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

	fmt.Println("=== ZIP entries ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Check document.xml paragraph count and structure
	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			
			paraCount := strings.Count(content, "<w:p>") + strings.Count(content, "<w:p ")
			fmt.Printf("\nTotal <w:p> in document.xml: %d\n", paraCount)
			
			// Count images
			imgCount := strings.Count(content, "<w:drawing>")
			fmt.Printf("Total <w:drawing> in document.xml: %d\n", imgCount)
			
			// Count tables
			tblCount := strings.Count(content, "<w:tbl>")
			fmt.Printf("Total <w:tbl> in document.xml: %d\n", tblCount)
			
			// Check for page breaks
			pageBreaks := strings.Count(content, `w:type="page"`)
			fmt.Printf("Total page breaks: %d\n", pageBreaks)
			
			// Check section properties
			sectPr := strings.Count(content, "<w:sectPr")
			fmt.Printf("Total <w:sectPr>: %d\n", sectPr)
			
			// Check header/footer references
			hdrRef := strings.Count(content, "headerReference")
			ftrRef := strings.Count(content, "footerReference")
			fmt.Printf("Header references: %d, Footer references: %d\n", hdrRef, ftrRef)
		}
		
		// Check headers
		if strings.HasPrefix(f.Name, "word/header") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			hasImg := strings.Contains(content, "<w:drawing>")
			fmt.Printf("\n%s: hasImage=%v len=%d\n", f.Name, hasImg, len(content))
			if len(content) < 2000 {
				fmt.Println(content)
			}
		}
		
		// Check footers
		if strings.HasPrefix(f.Name, "word/footer") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n%s: len=%d\n", f.Name, len(data))
			fmt.Println(string(data))
		}
	}
}
