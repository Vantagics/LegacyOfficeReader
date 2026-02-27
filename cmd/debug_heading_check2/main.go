package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find all Heading style references
			count := strings.Count(content, "Heading")
			fmt.Printf("Total 'Heading' occurrences: %d\n", count)
			
			// Find first heading
			idx := strings.Index(content, "Heading")
			if idx >= 0 {
				start := idx - 100
				if start < 0 { start = 0 }
				end := idx + 200
				if end > len(content) { end = len(content) }
				fmt.Printf("First Heading context:\n%s\n", content[start:end])
			}
			
			// Search for "引言" directly
			idx = strings.Index(content, "引言")
			if idx >= 0 {
				start := idx - 200
				if start < 0 { start = 0 }
				end := idx + 100
				if end > len(content) { end = len(content) }
				fmt.Printf("\n'引言' context:\n%s\n", content[start:end])
			}
		}
	}
}
