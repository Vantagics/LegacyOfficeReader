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
		fmt.Println("Error:", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				fmt.Println("Error:", err)
				os.Exit(1)
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			s := string(data)
			// Find paragraphs with numPr containing numId="2" (bullet list)
			count := strings.Count(s, `<w:numId w:val="2"/>`)
			fmt.Printf("Bullet list paragraphs (numId=2): %d\n", count)
			// Show a few examples
			parts := strings.Split(s, "<w:p>")
			shown := 0
			for i, part := range parts {
				if strings.Contains(part, `<w:numId w:val="2"/>`) && shown < 5 {
					// Extract text
					textStart := strings.Index(part, "<w:t")
					text := ""
					if textStart >= 0 {
						rest := part[textStart:]
						gt := strings.Index(rest, ">")
						if gt >= 0 {
							rest = rest[gt+1:]
							end := strings.Index(rest, "</w:t>")
							if end >= 0 {
								text = rest[:end]
							}
						}
					}
					// Extract ilvl
					ilvlStart := strings.Index(part, `<w:ilvl w:val="`)
					ilvl := "?"
					if ilvlStart >= 0 {
						rest := part[ilvlStart+15:]
						end := strings.Index(rest, `"`)
						if end >= 0 {
							ilvl = rest[:end]
						}
					}
					fmt.Printf("  P[%d] ilvl=%s text=%q\n", i-1, ilvl, text)
					shown++
				}
			}
			return
		}
	}
}
