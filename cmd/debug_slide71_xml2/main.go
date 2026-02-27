package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test_compare.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide71.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Find the large text box shape (安装Agent引流)
			idx := strings.Index(content, "11301412")
			if idx >= 0 {
				start := idx - 200
				if start < 0 { start = 0 }
				end := idx + 1000
				if end > len(content) { end = len(content) }
				fmt.Printf("=== Large text box shape ===\n%s\n", content[start:end])
			}
		}
	}
}
