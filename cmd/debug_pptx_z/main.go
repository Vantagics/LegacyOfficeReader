package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "ppt/slides/slide13.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Find "z" text
			idx := 0
			for {
				pos := strings.Index(content[idx:], ">z<")
				if pos < 0 {
					break
				}
				absPos := idx + pos
				start := absPos - 500
				if start < 0 {
					start = 0
				}
				end := absPos + 100
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("Found 'z' at pos %d:\n%s\n\n", absPos, content[start:end])
				idx = absPos + 3
			}
		}
	}
}
