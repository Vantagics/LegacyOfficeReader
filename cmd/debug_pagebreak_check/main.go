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

			// Find all page breaks
			idx := 0
			count := 0
			for {
				pos := strings.Index(content[idx:], `w:type="page"`)
				if pos < 0 { break }
				pos += idx
				count++
				
				// Find the next paragraph after this break
				nextP := strings.Index(content[pos:], "<w:p>")
				if nextP >= 0 {
					nextP += pos
					// Get the text of the next paragraph
					textStart := strings.Index(content[nextP:], "<w:t")
					if textStart >= 0 {
						textStart += nextP
						textEnd := strings.Index(content[textStart:], "</w:t>")
						if textEnd >= 0 {
							text := content[textStart:textStart+textEnd+6]
							if len(text) > 100 { text = text[:100] }
							fmt.Printf("Page break %d → next text: %s\n", count, text)
						}
					}
				}
				idx = pos + 10
			}
			fmt.Printf("\nTotal page breaks: %d\n", count)
		}
	}
}
