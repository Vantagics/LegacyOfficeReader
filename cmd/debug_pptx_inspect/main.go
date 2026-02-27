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

	// Print key XML files
	files := []string{
		"ppt/slideMasters/slideMaster1.xml",
		"ppt/slideMasters/_rels/slideMaster1.xml.rels",
		"ppt/slideLayouts/slideLayout4.xml",
		"ppt/slideLayouts/_rels/slideLayout4.xml.rels",
		"ppt/slides/_rels/slide4.xml.rels",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
	}

	for _, name := range files {
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				// Truncate if too long
				if len(content) > 3000 {
					content = content[:3000] + "\n... [truncated]"
				}
				fmt.Printf("=== %s ===\n%s\n\n", name, content)
			}
		}
	}

	// Also check slide 4 for a sample
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Just show first 2000 chars
			if len(content) > 2000 {
				content = content[:2000] + "\n... [truncated]"
			}
			fmt.Printf("=== slide4.xml (first 2000 chars) ===\n%s\n\n", content)
		}
	}

	// Check layout 4 for image shapes
	for _, f := range zr.File {
		if f.Name == "ppt/slideLayouts/slideLayout4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Find all r:embed references
			idx := 0
			for {
				pos := strings.Index(content[idx:], `r:embed="`)
				if pos < 0 {
					break
				}
				start := idx + pos + 9
				end := strings.Index(content[start:], `"`)
				if end < 0 {
					break
				}
				fmt.Printf("Layout4 embed: %s\n", content[start:start+end])
				idx = start + end + 1
			}
		}
	}
}
