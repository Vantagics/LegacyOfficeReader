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

	// Check context around each 0x01
	for i, r := range runes {
		if r == 0x01 {
			start := i - 20
			if start < 0 {
				start = 0
			}
			end := i + 20
			if end > len(runes) {
				end = len(runes)
			}
			// Show context with special chars marked
			fmt.Printf("0x01 at %d: ", i)
			for j := start; j < end; j++ {
				c := runes[j]
				if j == i {
					fmt.Print("[*IMG*]")
				} else if c == 0x13 {
					fmt.Print("[FB]")
				} else if c == 0x14 {
					fmt.Print("[FS]")
				} else if c == 0x15 {
					fmt.Print("[FE]")
				} else if c == 0x08 {
					fmt.Print("[OBJ]")
				} else if c == '\r' {
					fmt.Print("[\\r]")
				} else if c < 0x20 {
					fmt.Printf("[%02X]", c)
				} else {
					fmt.Printf("%c", c)
				}
			}
			fmt.Println()
		}
	}
}
