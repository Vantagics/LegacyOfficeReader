package main

import (
	"archive/zip"
	"fmt"
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

	for _, f := range r.File {
		if f.Name != "ppt/slides/slide26.xml" {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			continue
		}
		buf := make([]byte, 10*1024*1024)
		n, _ := rc.Read(buf)
		content := string(buf[:n])
		rc.Close()

		// Count shapes
		shapeCount := strings.Count(content, "<p:sp>")
		fmt.Printf("Slide 26: %d shapes in PPTX\n", shapeCount)
		
		// Count all unique fill colors
		idx := 0
		for {
			pos := strings.Index(content[idx:], `<a:solidFill><a:srgbClr val="`)
			if pos < 0 {
				break
			}
			start := idx + pos + len(`<a:solidFill><a:srgbClr val="`)
			end := start + 6
			if end <= len(content) {
				color := content[start:end]
				fmt.Printf("  Fill: %s\n", color)
			}
			idx = start + 1
		}
	}
}
