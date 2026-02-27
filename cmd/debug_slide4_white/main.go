package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name != "ppt/slides/slide4.xml" {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Find "监管要求日趋严格"
		idx := strings.Index(content, "监管要求日趋严格")
		if idx >= 0 {
			start := idx - 500
			if start < 0 { start = 0 }
			end := idx + 200
			if end > len(content) { end = len(content) }
			fmt.Printf("Context: %s\n", content[start:end])
		}

		// Also find the title shape
		idx = strings.Index(content, "背景：数据安全法规密集出台")
		if idx >= 0 {
			start := idx - 500
			if start < 0 { start = 0 }
			end := idx + 200
			if end > len(content) { end = len(content) }
			fmt.Printf("\nTitle context: %s\n", content[start:end])
		}
	}
}
