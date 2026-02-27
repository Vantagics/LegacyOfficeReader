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

			// Search for 目 and 录 separately
			idx1 := strings.Index(content, "目")
			if idx1 >= 0 {
				end := idx1 + 30
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("Found 目 at %d: %q\n", idx1, content[idx1:end])
				fmt.Printf("  bytes: %x\n", []byte(content[idx1:end]))
			}

			// Search for the TOC heading area
			idx2 := strings.Index(content, "录")
			if idx2 >= 0 {
				start := idx2 - 20
				if start < 0 {
					start = 0
				}
				end := idx2 + 10
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("Found 录 at %d: %q\n", idx2, content[start:end])
			}

			// Check for the specific text with ideographic space
			if strings.Contains(content, "目\u3000\u3000录") {
				fmt.Println("Found: 目　　录 (with ideographic spaces)")
			} else if strings.Contains(content, "目  录") {
				fmt.Println("Found: 目  录 (with regular spaces)")
			} else if strings.Contains(content, "目 录") {
				fmt.Println("Found: 目 录 (with single space)")
			}

			// Search for sz=44 bold text (the TOC heading style)
			idx3 := strings.Index(content, `w:val="44"`)
			if idx3 >= 0 {
				// Find the surrounding paragraph
				start := idx3 - 200
				if start < 0 {
					start = 0
				}
				end := idx3 + 200
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("\nAround sz=44: ...%q...\n", content[idx3-50:idx3+100])
			}
		}
	}
}
