package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "parse: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("no formatted content")
		return
	}

	dupCount := 0
	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.TrimSpace(text)
		runes := []rune(text)

		if len(runes) < 100 {
			continue
		}

		// Check if text contains a duplicate: try different split points
		half := len(runes) / 2
		// Check if the second half starts the same as the beginning
		checkLen := 30
		if half < checkLen {
			continue
		}
		first := string(runes[:checkLen])
		mid := string(runes[half : half+checkLen])
		if first == mid {
			dupCount++
			display := string(runes[:60])
			fmt.Printf("[%3d] DUP len=%d (half=%d) text=%q...\n", i+1, len(runes), half, display)
		}
	}

	fmt.Printf("\nTotal duplicates found: %d\n", dupCount)
}
