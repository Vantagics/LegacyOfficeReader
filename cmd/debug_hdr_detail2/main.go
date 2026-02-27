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

	fmt.Printf("Headers: %d\n", len(fc.Headers))
	for i, h := range fc.Headers {
		fmt.Printf("  H[%d]: %q\n", i, h)
	}
	fmt.Printf("Footers: %d\n", len(fc.Footers))
	for i, f := range fc.Footers {
		fmt.Printf("  F[%d]: %q\n", i, f)
	}
	fmt.Printf("HeadersRaw: %d\n", len(fc.HeadersRaw))
	for i, h := range fc.HeadersRaw {
		fmt.Printf("  HR[%d]: %q\n", i, h)
	}
	fmt.Printf("FootersRaw: %d\n", len(fc.FootersRaw))
	for i, f := range fc.FootersRaw {
		fmt.Printf("  FR[%d]: %q\n", i, f)
	}

	// Also check the full text to see what's in the header area
	text := d.GetText()
	fmt.Printf("\nFull text length: %d\n", len([]rune(text)))
	
	// Show first 200 chars
	runes := []rune(text)
	if len(runes) > 200 {
		runes = runes[:200]
	}
	fmt.Printf("First 200 chars: %q\n", string(runes))
	
	// Show last 200 chars
	runes2 := []rune(text)
	if len(runes2) > 200 {
		runes2 = runes2[len(runes2)-200:]
	}
	fmt.Printf("Last 200 chars: %q\n", string(runes2))

	// Check for section breaks
	for i, p := range fc.Paragraphs {
		if p.IsSectionBreak {
			ptext := ""
			for _, r := range p.Runs {
				ptext += r.Text
			}
			ptext = strings.TrimRight(ptext, "\r\n")
			fmt.Printf("Section break at P[%d]: type=%d text=%q\n", i, p.SectionType, ptext)
		}
	}
}
