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

			// Find all sectPr occurrences
			idx := 0
			count := 0
			for {
				pos := strings.Index(content[idx:], "<w:sectPr")
				if pos < 0 {
					break
				}
				pos += idx
				// Find end of sectPr
				end := strings.Index(content[pos:], "</w:sectPr>")
				if end < 0 {
					break
				}
				end += pos + 11
				
				// Show context around the sectPr
				contextStart := pos - 200
				if contextStart < 0 {
					contextStart = 0
				}
				
				fmt.Printf("=== sectPr[%d] at byte %d ===\n", count, pos)
				fmt.Printf("Content: %s\n", content[pos:end])
				
				// Show what's before it
				before := content[contextStart:pos]
				// Find last paragraph text before this sectPr
				lastP := strings.LastIndex(before, "</w:t>")
				if lastP >= 0 {
					textStart := strings.LastIndex(before[:lastP], ">")
					if textStart >= 0 {
						fmt.Printf("Text before: %q\n", before[textStart+1:lastP])
					}
				}
				fmt.Println()
				
				count++
				idx = end
			}
			
			// Count paragraphs
			pCount := strings.Count(content, "<w:p>") + strings.Count(content, "<w:p ")
			fmt.Printf("Total paragraphs: %d\n", pCount)
			
			// Check first few paragraphs
			fmt.Println("\n=== First 3 paragraphs ===")
			pIdx := 0
			searchIdx := 0
			for pIdx < 3 {
				pos := strings.Index(content[searchIdx:], "<w:p>")
				if pos < 0 {
					pos = strings.Index(content[searchIdx:], "<w:p ")
				}
				if pos < 0 {
					break
				}
				pos += searchIdx
				pEnd := strings.Index(content[pos:], "</w:p>")
				if pEnd < 0 {
					break
				}
				pEnd += pos + 6
				
				pContent := content[pos:pEnd]
				if len(pContent) > 300 {
					pContent = pContent[:300] + "..."
				}
				fmt.Printf("P[%d]: %s\n\n", pIdx, pContent)
				
				searchIdx = pEnd
				pIdx++
			}
		}
	}
}
