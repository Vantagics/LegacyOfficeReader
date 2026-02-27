package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Count paragraphs
			pCount := strings.Count(content, "<w:p>") + strings.Count(content, "<w:p ")
			fmt.Printf("Paragraph count in docx: %d\n", pCount)

			// Check for section properties
			sectCount := strings.Count(content, "<w:sectPr")
			fmt.Printf("Section properties count: %d\n", sectCount)

			// Check for page breaks
			pbCount := strings.Count(content, `w:type="page"`)
			fmt.Printf("Page breaks: %d\n", pbCount)

			// Check for tables
			tblCount := strings.Count(content, "<w:tbl>")
			fmt.Printf("Tables: %d\n", tblCount)

			// Check for images
			imgCount := strings.Count(content, "<pic:pic>")
			fmt.Printf("Images in document.xml: %d\n", imgCount)

			// Check for numbering references
			numCount := strings.Count(content, "<w:numId")
			fmt.Printf("Numbering references: %d\n", numCount)

			// Check for heading styles
			for i := 1; i <= 3; i++ {
				hCount := strings.Count(content, fmt.Sprintf(`w:val="Heading%d"`, i))
				fmt.Printf("Heading%d references: %d\n", i, hCount)
			}

			// Check for TOC styles
			for i := 1; i <= 3; i++ {
				tocCount := strings.Count(content, fmt.Sprintf(`w:val="TOC%d"`, i))
				fmt.Printf("TOC%d references: %d\n", i, tocCount)
			}

			// Check for text boxes
			txbxCount := strings.Count(content, "<wps:wsp>")
			fmt.Printf("Text boxes (wsp): %d\n", txbxCount)

			// Check for section break in sectPr
			if strings.Contains(content, `<w:type w:val="nextPage"`) {
				fmt.Println("Has nextPage section break")
			}
			if strings.Contains(content, `<w:headerReference`) {
				fmt.Println("Has header references in sectPr")
			}
			if strings.Contains(content, `<w:footerReference`) {
				fmt.Println("Has footer references in sectPr")
			}

			// Check page size
			if strings.Contains(content, `<w:pgSz`) {
				idx := strings.Index(content, `<w:pgSz`)
				end := idx + 100
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("Page size: %s\n", content[idx:end])
			}

			// Check margins
			if strings.Contains(content, `<w:pgMar`) {
				idx := strings.Index(content, `<w:pgMar`)
				end := idx + 150
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("Page margins: %s\n", content[idx:end])
			}

			// Check for titlePg
			if strings.Contains(content, `<w:titlePg`) {
				fmt.Println("Has titlePg (different first page)")
			}

			// Extract and show the sectPr content
			sectIdx := strings.LastIndex(content, "<w:sectPr")
			if sectIdx >= 0 {
				sectEnd := strings.Index(content[sectIdx:], "</w:sectPr>")
				if sectEnd >= 0 {
					fmt.Printf("\nFinal sectPr:\n%s\n", content[sectIdx:sectIdx+sectEnd+11])
				}
			}
		}
	}
}
