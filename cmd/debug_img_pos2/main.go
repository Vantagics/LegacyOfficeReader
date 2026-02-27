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

	// GetText() returns only main body text (truncated at ccpText)
	text := d.GetText()
	runes := []rune(text)
	img01 := 0
	obj08 := 0
	for i, r := range runes {
		if r == 0x01 {
			img01++
			fmt.Printf("  0x01 at rune position %d\n", i)
		}
		if r == 0x08 {
			obj08++
			fmt.Printf("  0x08 at rune position %d\n", i)
		}
	}
	fmt.Printf("Main body text length: %d runes\n", len(runes))
	fmt.Printf("0x01 (inline image) count: %d\n", img01)
	fmt.Printf("0x08 (drawn object) count: %d\n", obj08)
}
