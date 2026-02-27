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
	fmt.Printf("Main text length: %d\n", len(runes))

	// Show all 0x08 positions and surrounding context
	for i, r := range runes {
		if r == 0x08 {
			start := i - 3
			if start < 0 {
				start = 0
			}
			end := i + 4
			if end > len(runes) {
				end = len(runes)
			}
			context := ""
			for j := start; j < end; j++ {
				c := runes[j]
				if c < 0x20 {
					context += fmt.Sprintf("[%02X]", c)
				} else {
					context += string(c)
				}
			}
			fmt.Printf("  0x08 at CP %d: ...%s...\n", i, context)
		}
	}

	// Check what's around CPs 9150-9170
	fmt.Println("\n=== Around CP 9150-9170 ===")
	for i := 9150; i < 9170 && i < len(runes); i++ {
		r := runes[i]
		if r < 0x20 {
			fmt.Printf("  CP %d: \\x%02X\n", i, r)
		} else {
			fmt.Printf("  CP %d: %c\n", i, r)
		}
	}
}
