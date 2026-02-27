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
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check all layouts
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== %s ===\n%s\n", f.Name, string(data))
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
			fmt.Printf("\n=== theme1.xml ===\n%s\n", content)
		}
	}

	// Check slide master
	for _, f := range zr.File {
		if f.Name == "ppt/slideMasters/slideMaster1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== slideMaster1.xml ===\n%s\n", string(data))
		}
	}

	// Check which slides reference which layouts
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/_rels/slide") && strings.HasSuffix(f.Name, ".xml.rels") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Extract layout reference
			idx := strings.Index(content, "slideLayout")
			if idx > 0 {
				end := idx + 30
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("%s -> %s\n", f.Name, content[idx:end])
			}
		}
	}
}
