package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		panic(err)
	}
	defer zr.Close()

	// Check specific slides for title color
	checkSlides := []int{1, 4, 5, 13, 21, 41, 51, 61, 71}
	for _, slideNum := range checkSlides {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Find first few text runs with their colors
				fmt.Printf("=== Slide %d ===\n", slideNum)
				// Extract first 5 srgbClr values
				idx := 0
				count := 0
				for count < 8 {
					pos := strings.Index(content[idx:], `<a:srgbClr val="`)
					if pos < 0 {
						break
					}
					start := idx + pos + 16
					end := strings.Index(content[start:], `"`)
					if end < 0 {
						break
					}
					color := content[start : start+end]

					// Find nearby text
					textStart := strings.Index(content[start:], `<a:t>`)
					textEnd := -1
					text := ""
					if textStart >= 0 && textStart < 500 {
						textStart += start + 5
						textEnd = strings.Index(content[textStart:], `</a:t>`)
						if textEnd >= 0 {
							text = content[textStart : textStart+textEnd]
							if len([]rune(text)) > 30 {
								text = string([]rune(text)[:30]) + "..."
							}
						}
					}

					fmt.Printf("  color=%s text=%q\n", color, text)
					idx = start + end + 1
					count++
				}
				fmt.Println()
			}
		}
	}
}
