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

	// List all files
	for _, f := range zr.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Print layout files
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			hasBg := strings.Contains(content, "<p:bg>")
			hasBlip := strings.Contains(content, "blipFill")
			shapeCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ") + strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
			fmt.Printf("\n%s: hasBg=%v hasBlip=%v shapes=%d size=%d\n", f.Name, hasBg, hasBlip, shapeCount, len(data))
			// Show first 500 chars
			if len(content) > 500 {
				content = content[:500]
			}
			fmt.Println(content)
		}
	}

	// Print slide1 rels
	for _, f := range zr.File {
		if f.Name == "ppt/slides/_rels/slide1.xml.rels" || f.Name == "ppt/slides/_rels/slide2.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n%s:\n%s\n", f.Name, string(data))
		}
	}

	// Print slide1 XML (first 1000 chars)
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide1.xml" || f.Name == "ppt/slides/slide2.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			fmt.Printf("\n%s (size=%d):\n", f.Name, len(data))
			if len(content) > 2000 {
				content = content[:2000]
			}
			fmt.Println(content)
		}
	}

	// Print presentation.xml
	for _, f := range zr.File {
		if f.Name == "ppt/presentation.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n%s:\n%s\n", f.Name, string(data))
		}
	}
}
