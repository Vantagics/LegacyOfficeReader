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
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	// Get raw text and compute CP positions for each paragraph
	rawText := d.GetText()
	rawText = strings.ReplaceAll(rawText, "\r\n", "\r")

	runes := []rune(rawText)
	paraIdx := 0
	start := 0

	for i, r := range runes {
		if r == '\r' || r == 0x07 {
			pText := string(runes[start:i])
			display := strings.Map(func(c rune) rune {
				if c < 0x20 && c != '\t' { return '·' }
				return c
			}, pText)
			if len([]rune(display)) > 40 {
				display = string([]rune(display)[:40]) + "..."
			}

			if paraIdx >= 135 && paraIdx <= 140 {
				fmt.Printf("P%d: cpStart=%d cpEnd=%d text=%q\n", paraIdx, start, i, display)
			}
			start = i + 1
			paraIdx++
		}
	}
}
