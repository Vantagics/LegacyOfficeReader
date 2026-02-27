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

	// Show first 30 characters with their CP positions and hex values
	fmt.Println("=== First 30 characters of main text ===")
	limit := 30
	if limit > len(runes) {
		limit = len(runes)
	}
	for i := 0; i < limit; i++ {
		r := runes[i]
		display := string(r)
		if r < 0x20 {
			display = fmt.Sprintf("\\x%02X", r)
		}
		fmt.Printf("  CP %3d: U+%04X %s\n", i, r, display)
	}

	// Show around CP 12, 15, 108
	fmt.Println("\n=== Around CP 12 ===")
	for i := 8; i < 20 && i < len(runes); i++ {
		r := runes[i]
		display := string(r)
		if r < 0x20 {
			display = fmt.Sprintf("\\x%02X", r)
		}
		fmt.Printf("  CP %3d: U+%04X %s\n", i, r, display)
	}

	fmt.Println("\n=== Around CP 108 ===")
	for i := 104; i < 115 && i < len(runes); i++ {
		r := runes[i]
		display := string(r)
		if r < 0x20 {
			display = fmt.Sprintf("\\x%02X", r)
		}
		fmt.Printf("  CP %3d: U+%04X %s\n", i, r, display)
	}
}
