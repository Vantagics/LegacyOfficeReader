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

	fmt.Println("=== DOCX contents ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Read document.xml and show first 3000 chars
	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()

			content := string(data)

			// Count images
			imgCount := strings.Count(content, "<w:drawing>")
			fmt.Printf("\n=== document.xml: %d images ===\n", imgCount)

			// Find image positions
			pos := 0
			for i := 0; i < imgCount; i++ {
				idx := strings.Index(content[pos:], "<w:drawing>")
				if idx < 0 {
					break
				}
				absPos := pos + idx
				// Find the r:embed attribute
				embedStart := strings.Index(content[absPos:], `r:embed="`)
				if embedStart >= 0 {
					embedStart += absPos + 9
					embedEnd := strings.Index(content[embedStart:], `"`)
					if embedEnd >= 0 {
						relID := content[embedStart : embedStart+embedEnd]
						// Find surrounding paragraph context
						pStart := strings.LastIndex(content[:absPos], "<w:p>")
						if pStart < 0 {
							pStart = absPos - 100
						}
						// Get text content near the image
						textStart := pStart
						textEnd := absPos + 200
						if textEnd > len(content) {
							textEnd = len(content)
						}
						snippet := content[textStart:textEnd]
						// Extract text between <w:t> tags
						texts := extractTexts(snippet)
						fmt.Printf("  Image %d: relID=%s texts=%v\n", i+1, relID, texts)
					}
				}
				pos = absPos + 1
			}

			// Show first 2000 chars
			if len(content) > 2000 {
				content = content[:2000]
			}
			fmt.Printf("\n=== First 2000 chars of document.xml ===\n%s\n", content)
		}
	}
}

func extractTexts(xml string) []string {
	var texts []string
	pos := 0
	for {
		start := strings.Index(xml[pos:], "<w:t")
		if start < 0 {
			break
		}
		start += pos
		// Find end of opening tag
		tagEnd := strings.Index(xml[start:], ">")
		if tagEnd < 0 {
			break
		}
		tagEnd += start + 1
		// Find closing tag
		closeTag := strings.Index(xml[tagEnd:], "</w:t>")
		if closeTag < 0 {
			break
		}
		closeTag += tagEnd
		text := xml[tagEnd:closeTag]
		if text != "" {
			texts = append(texts, text)
		}
		pos = closeTag + 6
	}
	return texts
}
