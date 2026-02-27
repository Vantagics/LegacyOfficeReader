package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/reference.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	slideCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideCount++
		}
	}
	fmt.Printf("Reference PPTX: %d slides\n", slideCount)

	// Check slide 1
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if len(content) > 4000 {
				content = content[:4000]
			}
			fmt.Printf("\n=== Reference slide1.xml ===\n%s\n", content)
		}
	}

	// Check theme
	for _, f := range zr.File {
		if f.Name == "ppt/theme/theme1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if len(content) > 2000 {
				content = content[:2000]
			}
			fmt.Printf("\n=== Reference theme1.xml ===\n%s\n", content)
		}
	}
}
