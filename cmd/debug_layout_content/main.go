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

	// Check layout files
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Count images (pic elements)
			picCount := strings.Count(content, "<p:pic>")
			spCount := strings.Count(content, "<p:sp>")
			fmt.Printf("%s: %d pics, %d shapes, len=%d\n", f.Name, picCount, spCount, len(content))
		}
		// Check slide-layout relationships
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/_rels/slideLayout") && strings.HasSuffix(f.Name, ".xml.rels") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			imgCount := strings.Count(content, "image/")
			fmt.Printf("  %s: %d image rels\n", f.Name, imgCount)
		}
	}

	// Check slide-layout references from slides
	fmt.Println("\n--- Slide → Layout mapping ---")
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "ppt/slides/_rels/slide") && strings.HasSuffix(f.Name, ".xml.rels") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Find slideLayout reference
			idx := strings.Index(content, "slideLayout")
			if idx >= 0 {
				end := strings.Index(content[idx:], "\"")
				if end > 0 {
					slideNum := strings.TrimPrefix(f.Name, "ppt/slides/_rels/slide")
					slideNum = strings.TrimSuffix(slideNum, ".xml.rels")
					fmt.Printf("  slide%s → %s\n", slideNum, content[idx:idx+end])
				}
			}
		}
	}
}
