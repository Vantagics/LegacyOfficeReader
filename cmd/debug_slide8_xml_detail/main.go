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

	for _, zf := range r.File {
		if zf.Name != "ppt/slides/slide8.xml" {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Find Shape 20 (the big F0F0F0 background freeform)
		// It should be around "Shape 20" in the XML
		targets := []string{"Shape 20", "Shape 22", "Shape 23"}
		for _, target := range targets {
			idx := strings.Index(content, target)
			if idx < 0 {
				fmt.Printf("%s: NOT FOUND\n\n", target)
				continue
			}
			// Show 500 chars after the shape name
			end := idx + 600
			if end > len(content) {
				end = len(content)
			}
			fmt.Printf("=== %s ===\n%s\n\n", target, content[idx:end])
		}
	}
}
