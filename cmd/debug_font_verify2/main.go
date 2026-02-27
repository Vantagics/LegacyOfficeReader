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
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find typeface attributes
			idx := 0
			for {
				pos := strings.Index(content[idx:], `typeface="`)
				if pos < 0 {
					break
				}
				pos += idx + 10
				end := strings.Index(content[pos:], `"`)
				if end < 0 {
					break
				}
				fontName := content[pos : pos+end]
				fmt.Printf("Font: %q (hex: %x)\n", fontName, []byte(fontName))
				idx = pos + end + 1
			}
			break
		}
	}

	// Also check theme fonts
	for _, f := range zr.File {
		if f.Name == "ppt/theme/theme1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			fmt.Println("\n--- Theme fonts ---")
			idx := 0
			for {
				pos := strings.Index(content[idx:], `typeface="`)
				if pos < 0 {
					break
				}
				pos += idx + 10
				end := strings.Index(content[pos:], `"`)
				if end < 0 {
					break
				}
				fontName := content[pos : pos+end]
				fmt.Printf("Font: %q (hex: %x)\n", fontName, []byte(fontName))
				idx = pos + end + 1
			}
			break
		}
	}
}
