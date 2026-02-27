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
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check slide 63 for text colors
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide63.xml" {
			rc, err := f.Open()
			if err != nil {
				fmt.Fprintf(os.Stderr, "Failed to open: %v\n", err)
				return
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Count FFFFFF colors
			whiteCount := strings.Count(content, `val="FFFFFF"`)
			blackCount := strings.Count(content, `val="000000"`)
			fmt.Printf("Slide 63: FFFFFF count=%d, 000000 count=%d\n", whiteCount, blackCount)

			// Find text with FFFFFF on E9EBF5 fill
			// Look for patterns like solidFill with E9EBF5 near solidFill with FFFFFF
			idx := 0
			e9Count := 0
			for {
				pos := strings.Index(content[idx:], `val="E9EBF5"`)
				if pos < 0 {
					break
				}
				idx += pos + 1
				e9Count++
			}
			fmt.Printf("E9EBF5 fill count=%d\n", e9Count)

			// Show a sample shape with E9EBF5 fill
			idx = 0
			shown := 0
			for shown < 3 {
				pos := strings.Index(content[idx:], `val="E9EBF5"`)
				if pos < 0 {
					break
				}
				absPos := idx + pos
				// Show context around this
				start := absPos - 200
				if start < 0 {
					start = 0
				}
				end := absPos + 500
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("\n--- Sample E9EBF5 shape (pos=%d) ---\n%s\n", absPos, content[start:end])
				idx = absPos + 1
				shown++
			}
			return
		}
	}
	fmt.Println("slide63.xml not found")
}
