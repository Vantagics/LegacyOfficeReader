package main

import (
	"fmt"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		return
	}

	rawText := d.GetText()
	runes := []rune(rawText)
	fmt.Printf("GetText() rune count: %d\n", len(runes))
	
	// Show last 20 chars
	if len(runes) > 20 {
		fmt.Printf("Last 20 chars:\n")
		for i := len(runes) - 20; i < len(runes); i++ {
			r := runes[i]
			if r < 0x20 {
				fmt.Printf("  [%d] [%02X]\n", i, r)
			} else {
				fmt.Printf("  [%d] U+%04X %c\n", i, r, r)
			}
		}
	}
}
