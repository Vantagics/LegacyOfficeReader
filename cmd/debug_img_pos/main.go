package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	// Find paragraphs that contain 0x01 image placeholders
	fmt.Println("=== Paragraphs with image placeholders (0x01) ===")
	imgCount := 0
	for i, p := range fc.Paragraphs {
		for j, r := range p.Runs {
			count := strings.Count(r.Text, "\x01")
			if count > 0 {
				fmt.Printf("P%d Run%d: %d images, text=%q\n", i, j, count, truncate(r.Text, 80))
				imgCount += count
			}
		}
	}
	fmt.Printf("Total image placeholders in formatted content: %d\n", imgCount)

	// Check raw text for 0x01 and 0x08
	text := d.GetText()
	runes := []rune(text)
	img01 := 0
	obj08 := 0
	for _, r := range runes {
		if r == 0x01 {
			img01++
		}
		if r == 0x08 {
			obj08++
		}
	}
	fmt.Printf("\nRaw text: 0x01 count=%d, 0x08 count=%d\n", img01, obj08)

	imgs := d.GetImages()
	fmt.Printf("Extracted images: %d\n", len(imgs))
}

func truncate(s string, maxLen int) string {
	runes := []rune(s)
	if len(runes) <= maxLen {
		return s
	}
	return string(runes[:maxLen]) + "..."
}
