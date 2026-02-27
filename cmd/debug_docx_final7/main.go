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

			// Find the actual heading "引言" (not in TOC)
			// Search for Heading1 style
			searches := []string{
				`Heading1"/>`,
				`Heading2"/>`,
				`Heading3"/>`,
				`<w:br w:type="page"/>`,
			}
			for _, s := range searches {
				idx := 0
				count := 0
				for {
					pos := strings.Index(content[idx:], s)
					if pos < 0 {
						break
					}
					idx += pos
					start := idx - 200
					if start < 0 {
						start = 0
					}
					end := idx + 300
					if end > len(content) {
						end = len(content)
					}
					count++
					if count <= 3 {
						fmt.Printf("=== %s (#%d) ===\n%s\n\n", s, count, content[start:end])
					}
					idx += len(s)
				}
				fmt.Printf("Total '%s': %d\n\n", s, count)
			}

			// Check for the deployment diagram images
			idx := strings.Index(content, "高级威胁检测及回溯方案")
			if idx > 0 {
				// Find the next drawing after this
				drawIdx := strings.Index(content[idx:], "<w:drawing>")
				if drawIdx > 0 {
					start := idx + drawIdx - 50
					end := start + 500
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("=== Drawing after '高级威胁检测及回溯方案' ===\n%s\n\n", content[start:end])
				}
			}
		}
	}
}
