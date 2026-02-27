package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer r.Close()

	// List all files in the DOCX
	fmt.Println("=== DOCX Contents ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Check document.xml.rels for image references
	for _, f := range r.File {
		if f.Name == "word/_rels/document.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			fmt.Println("\n=== Document Relationships ===")
			// Count image relationships
			imgCount := strings.Count(content, "image/")
			fmt.Printf("Image relationships: %d\n", imgCount)
			// Show all relationships
			parts := strings.Split(content, "<Relationship ")
			for _, part := range parts[1:] {
				end := strings.Index(part, "/>")
				if end > 0 {
					fmt.Printf("  %s\n", part[:end])
				}
			}
		}
	}

	// Check [Content_Types].xml
	for _, f := range r.File {
		if f.Name == "[Content_Types].xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			fmt.Println("\n=== Content Types ===")
			// Show overrides
			parts := strings.Split(content, "<Override ")
			for _, part := range parts[1:] {
				end := strings.Index(part, "/>")
				if end > 0 {
					fmt.Printf("  %s\n", part[:end])
				}
			}
		}
	}
}
