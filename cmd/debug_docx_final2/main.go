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
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Count various elements
			fmt.Printf("Total <w:p> tags: %d\n", strings.Count(content, "<w:p>") + strings.Count(content, "<w:p "))
			fmt.Printf("Total <w:drawing> tags: %d\n", strings.Count(content, "<w:drawing>"))
			fmt.Printf("Total <w:tbl> tags: %d\n", strings.Count(content, "<w:tbl>"))
			fmt.Printf("Total <w:sectPr> tags: %d\n", strings.Count(content, "<w:sectPr>"))
			fmt.Printf("Total <w:br w:type=\"page\"/> tags: %d\n", strings.Count(content, `<w:br w:type="page"/>`))
			fmt.Printf("Total <wp:inline> tags: %d\n", strings.Count(content, "<wp:inline"))
			fmt.Printf("Total <wp:anchor> tags: %d\n", strings.Count(content, "<wp:anchor"))

			// Find all image references
			idx := 0
			for {
				pos := strings.Index(content[idx:], `r:embed="rImg`)
				if pos < 0 {
					break
				}
				idx += pos
				end := strings.Index(content[idx:], `"`)
				if end < 0 {
					break
				}
				start2 := strings.Index(content[idx:], `"`)
				end2 := strings.Index(content[idx+start2+1:], `"`)
				relID := content[idx+start2+1 : idx+start2+1+end2]
				fmt.Printf("Image ref: %s\n", relID)
				idx += end + 1
			}

			// Show the last 2000 chars (section properties)
			if len(content) > 2000 {
				fmt.Printf("\n=== Last 2000 chars ===\n%s\n", content[len(content)-2000:])
			}

			// Check for TOC
			fmt.Printf("\nTOC begin: %d\n", strings.Count(content, `TOC \o`))
			fmt.Printf("TOC style refs: TOC1=%d TOC2=%d TOC3=%d\n",
				strings.Count(content, `"TOC1"`),
				strings.Count(content, `"TOC2"`),
				strings.Count(content, `"TOC3"`))

			// Check heading styles
			for i := 1; i <= 3; i++ {
				fmt.Printf("Heading%d refs: %d\n", i, strings.Count(content, fmt.Sprintf(`"Heading%d"`, i)))
			}

			// Check for numbering
			fmt.Printf("numPr refs: %d\n", strings.Count(content, "<w:numPr>"))
		}
	}
}
