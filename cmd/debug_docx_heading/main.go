package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	path := "testfie/test.docx"
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

			// Find Heading1 paragraphs
			idx := 0
			for {
				next := strings.Index(content[idx:], `Heading1`)
				if next < 0 {
					break
				}
				idx += next
				// Find enclosing <w:p>
				pStart := strings.LastIndex(content[:idx], "<w:p>")
				if pStart < 0 {
					pStart = strings.LastIndex(content[:idx], "<w:p ")
				}
				pEnd := strings.Index(content[idx:], "</w:p>")
				if pStart >= 0 && pEnd >= 0 {
					para := content[pStart : idx+pEnd+6]
					if len(para) > 600 {
						para = para[:600] + "..."
					}
					fmt.Printf("%s\n\n", para)
				}
				idx += 8
			}
		}
	}
}
