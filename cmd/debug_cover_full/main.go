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

			// Find the first 15 <w:p> elements (cover page area)
			parts := strings.Split(content, "<w:p>")
			for i := 1; i <= 15 && i < len(parts); i++ {
				part := parts[i]
				end := strings.Index(part, "</w:p>")
				if end > 0 {
					part = part[:end]
				}
				fmt.Printf("\n=== P[%d] (len=%d) ===\n%s\n", i, len(part), part)
			}
		}
	}
}
