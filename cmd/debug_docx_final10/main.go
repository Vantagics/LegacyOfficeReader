package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "word/numbering.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("=== numbering.xml ===\n%s\n", string(data))
		}
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find heading paragraphs and show their numPr
			idx := strings.Index(content, `Heading1"/>`)
			if idx > 0 {
				start := idx - 100
				end := idx + 200
				if start < 0 { start = 0 }
				if end > len(content) { end = len(content) }
				fmt.Printf("\n=== First Heading1 numPr ===\n%s\n", content[start:end])
			}

			// Count numId=1 and numId=2 references
			fmt.Printf("\nnumId val=\"1\": %d\n", strings.Count(content, `numId w:val="1"`))
			fmt.Printf("numId val=\"2\": %d\n", strings.Count(content, `numId w:val="2"`))
		}
	}
}
