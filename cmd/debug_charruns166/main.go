package main

import (
	"fmt"
	"os"

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

	// Access charRuns via DebugRanges or directly
	// We need to find the CP range for paragraph 166
	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("no formatted content")
		return
	}

	// Calculate CP position of paragraph 166
	text := d.GetText()
	runes := []rune(text)

	// Split by paragraph separators to find CP of paragraph 166
	cpPos := uint32(0)
	paraIdx := 0
	start := 0
	for i, r := range runes {
		if r == '\r' || r == 0x07 {
			if paraIdx == 165 { // 0-indexed
				fmt.Printf("Paragraph 166: CP range [%d, %d]\n", cpPos, cpPos+uint32(i-start))
				fmt.Printf("Text (first 100 chars): %q\n\n", string(runes[start:min(start+100, i)]))
				break
			}
			cpPos = uint32(i + 1)
			start = i + 1
			paraIdx++
		}
	}

	// Now print charRuns that overlap this range
	fmt.Println("CharRuns overlapping this paragraph:")
	d.DebugRanges()
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
