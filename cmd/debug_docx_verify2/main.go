package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	path := "testfie/test_v2.docx"
	if len(os.Args) > 1 {
		path = os.Args[1]
	}

	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	fmt.Println("=== DOCX Contents ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Read and analyze key files
	for _, name := range []string{"word/document.xml", "word/footer1.xml", "word/header1.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"} {
		for _, f := range r.File {
			if f.Name == name {
				rc, err := f.Open()
				if err != nil {
					continue
				}
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				fmt.Printf("\n=== %s ===\n", name)
				if name == "word/document.xml" {
					// Check for sectPr elements
					count := strings.Count(content, "<w:sectPr")
					fmt.Printf("  sectPr count: %d\n", count)

					// Check for footer references
					ftrCount := strings.Count(content, "footerReference")
					fmt.Printf("  footerReference count: %d\n", ftrCount)

					// Check for header references
					hdrCount := strings.Count(content, "headerReference")
					fmt.Printf("  headerReference count: %d\n", hdrCount)

					// Check first 500 chars of body
					bodyIdx := strings.Index(content, "<w:body>")
					if bodyIdx >= 0 {
						end := bodyIdx + 800
						if end > len(content) {
							end = len(content)
						}
						fmt.Printf("  First body content:\n%s\n", content[bodyIdx:end])
					}

					// Find all sectPr blocks
					idx := 0
					sectNum := 0
					for {
						pos := strings.Index(content[idx:], "<w:sectPr")
						if pos < 0 {
							break
						}
						sectNum++
						start := idx + pos
						endPos := strings.Index(content[start:], "</w:sectPr>")
						if endPos < 0 {
							break
						}
						sectContent := content[start : start+endPos+len("</w:sectPr>")]
						fmt.Printf("\n  sectPr #%d:\n%s\n", sectNum, sectContent)
						idx = start + endPos + len("</w:sectPr>")
					}
				} else {
					// Print full content for smaller files
					if len(content) > 2000 {
						fmt.Printf("%s...\n", content[:2000])
					} else {
						fmt.Println(content)
					}
				}
			}
		}
	}
}
