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

			// Find "C-创建" or "创建，A" 
			if strings.Contains(content, "创建") {
				// Find all occurrences
				idx := 0
				for {
					pos := strings.Index(content[idx:], "创建")
					if pos < 0 {
						break
					}
					pos += idx
					start := pos - 30
					if start < 0 {
						start = 0
					}
					end := pos + 30
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("Found 创建 at %d: %q\n", pos, content[start:end])
					idx = pos + 3
				}
			}
		}
	}
}
