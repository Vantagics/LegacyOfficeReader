package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	path := "testfie/test_v3.docx"
	if len(os.Args) > 1 {
		path = os.Args[1]
	}

	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Count inline vs anchor images
			inlineCount := strings.Count(content, "<wp:inline")
			anchorCount := strings.Count(content, "<wp:anchor")
			fmt.Printf("Inline images: %d\n", inlineCount)
			fmt.Printf("Anchor images: %d\n", anchorCount)
			fmt.Printf("Total drawings: %d\n", strings.Count(content, "<w:drawing>"))

			// Count unique image references
			imgRefs := make(map[string]int)
			idx := 0
			for {
				pos := strings.Index(content[idx:], `r:embed="`)
				if pos < 0 {
					break
				}
				start := idx + pos + len(`r:embed="`)
				end := strings.Index(content[start:], `"`)
				if end < 0 {
					break
				}
				ref := content[start : start+end]
				imgRefs[ref]++
				idx = start + end
			}
			fmt.Printf("\nImage references:\n")
			for ref, count := range imgRefs {
				fmt.Printf("  %s: %d times\n", ref, count)
			}

			// Check for page breaks
			fmt.Printf("\nPage breaks (br type=page): %d\n", strings.Count(content, `w:type="page"`))
			fmt.Printf("Page break before: %d\n", strings.Count(content, `<w:pageBreakBefore/>`))

			// Check section breaks
			fmt.Printf("Section breaks (sectPr in pPr): %d\n", strings.Count(content, `<w:pPr><w:sectPr`) + strings.Count(content, `</w:jc><w:sectPr`))
			fmt.Printf("Final sectPr: %d\n", strings.Count(content, `</w:sectPr>`))
		}
	}
}
