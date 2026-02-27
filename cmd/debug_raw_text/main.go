package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	text := d.GetText()
	runes := []rune(text)

	// Show first 500 characters with hex codes for special chars
	limit := 500
	if len(runes) < limit {
		limit = len(runes)
	}

	fmt.Printf("=== Raw text first %d runes ===\n", limit)
	for i := 0; i < limit; i++ {
		r := runes[i]
		if r == '\r' {
			fmt.Printf("[\\r]")
		} else if r == '\n' {
			fmt.Printf("[\\n]")
		} else if r == '\t' {
			fmt.Printf("[\\t]")
		} else if r == 0x01 {
			fmt.Printf("[IMG]")
		} else if r == 0x07 {
			fmt.Printf("[CELL]")
		} else if r == 0x08 {
			fmt.Printf("[OBJ]")
		} else if r == 0x0C {
			fmt.Printf("[PB]")
		} else if r == 0x13 {
			fmt.Printf("[FBEGIN]")
		} else if r == 0x14 {
			fmt.Printf("[FSEP]")
		} else if r == 0x15 {
			fmt.Printf("[FEND]")
		} else if r < 0x20 {
			fmt.Printf("[0x%02X]", r)
		} else {
			fmt.Printf("%c", r)
		}
	}
	fmt.Println()

	// Count special chars
	imgCount := 0
	for _, r := range runes {
		if r == 0x01 {
			imgCount++
		}
	}
	fmt.Printf("\nTotal runes: %d, Image placeholders (0x01): %d\n", len(runes), imgCount)

	// Show images
	imgs := d.GetImages()
	fmt.Printf("Images extracted: %d\n", len(imgs))
	for i, img := range imgs {
		fmt.Printf("  Image %d: format=%d, size=%d bytes\n", i, img.Format, len(img.Data))
	}
}
