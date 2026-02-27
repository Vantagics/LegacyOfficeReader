package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	// Check the converted test.docx
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
			fmt.Printf("Total <w:p> elements: %d\n", pCount)

			// Count tables
			tCount := strings.Count(content, "<w:tbl>")
			fmt.Printf("Total <w:tbl> elements: %d\n", tCount)

			// Count images
			imgCount := strings.Count(content, "<w:drawing>")
			fmt.Printf("Total <w:drawing> elements: %d\n", imgCount)

			// Count page breaks
			pbCount := strings.Count(content, `w:type="page"`)
			fmt.Printf("Total page breaks: %d\n", pbCount)

			// Count section breaks
			sbCount := strings.Count(content, "<w:sectPr>")
			fmt.Printf("Total <w:sectPr> elements: %d\n", sbCount)

			// Count headings
			for i := 1; i <= 3; i++ {
				hCount := strings.Count(content, fmt.Sprintf(`w:val="Heading%d"`, i))
				fmt.Printf("Heading%d count: %d\n", i, hCount)
			}

			// Show paragraphs around page breaks
			fmt.Println("\n=== Content around page breaks ===")
			idx := 0
			for {
				pos := strings.Index(content[idx:], `w:type="page"`)
				if pos == -1 {
					break
				}
				absPos := idx + pos
				start := absPos - 200
				if start < 0 {
					start = 0
				}
				end := absPos + 200
				if end > len(content) {
					end = len(content)
				}
				snippet := content[start:end]
				snippet = strings.ReplaceAll(snippet, "<w:p>", "\n  <w:p>")
				fmt.Printf("--- Page break at offset %d ---\n%s\n", absPos, snippet)
				idx = absPos + 13
			}
		}
	}
}
