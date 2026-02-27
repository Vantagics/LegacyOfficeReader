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
				fmt.Fprintf(os.Stderr, "Error opening: %v\n", err)
				return
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find all eastAsia font references
			idx := 0
			count := 0
			for {
				pos := strings.Index(content[idx:], "eastAsia=")
				if pos < 0 {
					break
				}
				pos += idx
				end := pos + 40
				if end > len(content) {
					end = len(content)
				}
				snippet := content[pos:end]
				fmt.Printf("eastAsia[%d]: %q\n", count, snippet)
				fmt.Printf("  raw bytes: %x\n", []byte(snippet))
				count++
				idx = pos + 10
				if count > 5 {
					break
				}
			}

			// Find first Chinese text
			for i, r := range content {
				if r >= 0x4E00 && r <= 0x9FFF {
					start := i - 20
					if start < 0 {
						start = 0
					}
					end := i + 40
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("\nFirst CJK char at byte %d: %q\n", i, content[start:end])
					break
				}
			}

			// Check styles.xml too
			fmt.Println("\n--- Checking styles.xml ---")
		}
		if f.Name == "word/styles.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			
			// Find eastAsia in styles
			idx := strings.Index(content, "eastAsia=")
			if idx >= 0 {
				end := idx + 50
				if end > len(content) {
					end = len(content)
				}
				snippet := content[idx:end]
				fmt.Printf("styles eastAsia: %q\n", snippet)
				fmt.Printf("  raw bytes: %x\n", []byte(snippet))
			}
			
			// Check if 宋体 appears
			if strings.Contains(content, "宋体") {
				fmt.Println("styles.xml contains 宋体 correctly")
			} else {
				fmt.Println("styles.xml does NOT contain 宋体")
				// Find what's there instead
				idx2 := strings.Index(content, `eastAsia="`)
				if idx2 >= 0 {
					end := idx2 + 30
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("  Found: %q\n", content[idx2:end])
				}
			}
		}
	}
}
