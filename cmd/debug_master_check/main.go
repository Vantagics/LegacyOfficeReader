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

	for _, f := range zr.File {
		if f.Name == "ppt/slideMasters/slideMaster1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Count layout refs
			count := strings.Count(content, "<p:sldLayoutId")
			fmt.Printf("Layout refs in master: %d\n", count)

			// Extract the sldLayoutIdLst section
			start := strings.Index(content, "<p:sldLayoutIdLst>")
			end := strings.Index(content, "</p:sldLayoutIdLst>")
			if start >= 0 && end >= 0 {
				fmt.Println(content[start : end+len("</p:sldLayoutIdLst>")])
			}
			break
		}
	}

	// Check master rels
	for _, f := range zr.File {
		if f.Name == "ppt/slideMasters/_rels/slideMaster1.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\nMaster rels:\n%s\n", string(data))
			break
		}
	}

	// List all layout files
	fmt.Println("\nLayout files:")
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/") {
			fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
		}
	}

	// Check content types for layouts
	for _, f := range zr.File {
		if f.Name == "[Content_Types].xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			layoutCount := strings.Count(content, "slideLayout")
			fmt.Printf("\nContent types layout entries: %d\n", layoutCount)
			break
		}
	}
}
