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

	fc := d.GetFormattedContent()
	
	// Count alignments
	alignCounts := map[uint8]int{}
	for _, p := range fc.Paragraphs {
		alignCounts[p.Props.Alignment]++
	}
	fmt.Printf("Alignment counts:\n")
	fmt.Printf("  0 (left): %d\n", alignCounts[0])
	fmt.Printf("  1 (center): %d\n", alignCounts[1])
	fmt.Printf("  2 (right): %d\n", alignCounts[2])
	fmt.Printf("  3 (both/justify): %d\n", alignCounts[3])

	// Show paragraphs with justify alignment
	fmt.Printf("\nJustified paragraphs:\n")
	for i, p := range fc.Paragraphs {
		if p.Props.Alignment == 3 {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			fmt.Printf("  P%d: %q\n", i, text)
		}
	}

	// Check line spacing
	spacingCounts := 0
	for _, p := range fc.Paragraphs {
		if p.Props.LineSpacing != 0 {
			spacingCounts++
		}
	}
	fmt.Printf("\nParagraphs with line spacing: %d\n", spacingCounts)

	// Check font sizes used
	fontSizes := map[uint16]int{}
	for _, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			if r.Props.FontSize > 0 {
				fontSizes[r.Props.FontSize]++
			}
		}
	}
	fmt.Printf("\nFont sizes used:\n")
	for sz, count := range fontSizes {
		fmt.Printf("  %d half-pts (%.1f pt): %d runs\n", sz, float64(sz)/2, count)
	}

	// Check fonts used
	fontNames := map[string]int{}
	for _, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			if r.Props.FontName != "" {
				fontNames[r.Props.FontName]++
			}
		}
	}
	fmt.Printf("\nFonts used:\n")
	for name, count := range fontNames {
		fmt.Printf("  %q: %d runs\n", name, count)
	}
}
