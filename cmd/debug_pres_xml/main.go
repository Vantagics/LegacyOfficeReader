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
		if f.Name == "ppt/presentation.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Extract sldIdLst section
			start := strings.Index(content, "<p:sldIdLst>")
			end := strings.Index(content, "</p:sldIdLst>")
			if start >= 0 && end >= 0 {
				section := content[start : end+len("</p:sldIdLst>")]
				// Count sldId entries
				count := strings.Count(section, "<p:sldId")
				fmt.Printf("sldId count: %d\n", count)

				// Show last few entries
				entries := strings.Split(section, "<p:sldId")
				if len(entries) > 3 {
					for i := len(entries) - 3; i < len(entries); i++ {
						fmt.Printf("  Entry %d: <p:sldId%s\n", i, entries[i])
					}
				}
			}
			break
		}
	}
}
