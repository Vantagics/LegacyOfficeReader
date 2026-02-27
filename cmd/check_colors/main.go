package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check for scheme-like colors (040000, 010000, etc.) in slide XMLs
	re := regexp.MustCompile(`val="([0-9A-Fa-f]{6})"`)
	badColors := map[string]int{}

	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()

			matches := re.FindAllStringSubmatch(string(data), -1)
			for _, m := range matches {
				color := m[1]
				// Check for suspicious scheme-like colors
				if color == "040000" || color == "010000" || color == "050000" || color == "060000" || color == "070000" || color == "080000" || color == "000004" || color == "000001" {
					badColors[color]++
				}
			}
		}
	}

	if len(badColors) > 0 {
		fmt.Println("Suspicious scheme-like colors found in PPTX:")
		for c, count := range badColors {
			fmt.Printf("  %s: %d occurrences\n", c, count)
		}
	} else {
		fmt.Println("No suspicious scheme-like colors found - scheme resolution working!")
	}

	// Also check slide 71 (the cloud deployment slide) for specific colors
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide71.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()

			content := string(data)
			matches := re.FindAllStringSubmatch(content, -1)
			colorCounts := map[string]int{}
			for _, m := range matches {
				colorCounts[m[1]]++
			}
			fmt.Printf("\nSlide 71 color distribution:\n")
			for c, count := range colorCounts {
				fmt.Printf("  %s: %d\n", c, count)
			}
		}
	}
}
