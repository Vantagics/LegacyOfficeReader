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
			
			// Find the first few <w:p> elements
			idx := 0
			count := 0
			for count < 15 {
				next := strings.Index(content[idx:], "<w:p>")
				if next < 0 {
					next = strings.Index(content[idx:], "<w:p ")
				}
				if next < 0 {
					break
				}
				idx += next
				// Find end of this paragraph
				end := strings.Index(content[idx:], "</w:p>")
				if end < 0 {
					break
				}
				para := content[idx : idx+end+6]
				if len(para) > 500 {
					para = para[:500] + "..."
				}
				fmt.Printf("Para[%d]: %s\n\n", count, para)
				idx += end + 6
				count++
			}
			
			// Find the sectPr at the end
			lastSect := strings.LastIndex(content, "<w:sectPr>")
			if lastSect >= 0 {
				fmt.Printf("\nFinal sectPr: %s\n", content[lastSect:])
			}
		}
	}
}
