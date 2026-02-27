package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		panic(err)
	}
	defer zr.Close()

	// Check slide 2 (CONTENTS page) for color issues
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide2.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("=== Slide 2 (CONTENTS) ===\n%s\n\n", string(data))
		}
	}

	// Check slide 5 for color issues (has 13 empty colors in source)
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide5.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Count color occurrences
			fmt.Printf("=== Slide 5 Analysis ===\n")
			fmt.Printf("  Contains 'val=\"000000\"': %d\n", strings.Count(content, `val="000000"`))
			fmt.Printf("  Contains 'val=\"FFFFFF\"': %d\n", strings.Count(content, `val="FFFFFF"`))
			fmt.Printf("  Contains 'val=\"003296\"': %d\n", strings.Count(content, `val="003296"`))
			fmt.Printf("  Contains 'val=\"\"': %d\n", strings.Count(content, `val=""`))
			// Show first 3000 chars
			if len(content) > 3000 {
				content = content[:3000]
			}
			fmt.Printf("  First 3000 chars:\n%s\n\n", content)
		}
	}

	// Check layout 4 (most used) for watermark
	for _, f := range zr.File {
		if f.Name == "ppt/slideLayouts/slideLayout4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("=== Layout 4 (most used) ===\n%s\n\n", string(data))
		}
	}

	// Check layout rels
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/_rels/") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("=== %s ===\n%s\n\n", f.Name, string(data))
		}
	}

	// Check presentation.xml for slide size
	for _, f := range zr.File {
		if f.Name == "ppt/presentation.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("=== presentation.xml ===\n%s\n\n", string(data))
		}
	}
}
