package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	path := "testfie/test_new2.docx"
	if len(os.Args) > 1 {
		path = os.Args[1]
	}

	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "failed to open: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, name := range []string{"word/header1.xml", "word/header2.xml", "word/header3.xml", "word/footer1.xml"} {
		content := readZip(r, name)
		fmt.Printf("=== %s ===\n", name)
		if len(content) > 1000 {
			fmt.Printf("  Length: %d bytes\n", len(content))
			// Check for images
			if strings.Contains(content, "wp:inline") || strings.Contains(content, "wp:anchor") {
				fmt.Printf("  Contains images: YES\n")
			}
			// Check for text
			if strings.Contains(content, "<w:t") {
				fmt.Printf("  Contains text: YES\n")
				// Extract text snippets
				start := 0
				for {
					idx := strings.Index(content[start:], "<w:t")
					if idx < 0 {
						break
					}
					tStart := start + idx
					tEnd := strings.Index(content[tStart:], "</w:t>")
					if tEnd < 0 {
						break
					}
					snippet := content[tStart : tStart+tEnd+6]
					if len(snippet) > 100 {
						snippet = snippet[:100] + "..."
					}
					fmt.Printf("    %s\n", snippet)
					start = tStart + tEnd + 6
				}
			} else {
				fmt.Printf("  Contains text: NO\n")
			}
		} else {
			fmt.Printf("  %s\n", content)
		}
		fmt.Println()
	}
}

func readZip(r *zip.ReadCloser, name string) string {
	for _, f := range r.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			return string(data)
		}
	}
	return ""
}
