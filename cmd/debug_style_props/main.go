package main

import (
	"fmt"
	"os"

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

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	// Check alignment distribution
	alignCounts := make(map[uint8]int)
	for _, p := range fc.Paragraphs {
		alignCounts[p.Props.Alignment]++
	}
	fmt.Println("Alignment distribution:")
	names := []string{"left(0)", "center(1)", "right(2)", "justify(3)"}
	for a := uint8(0); a <= 3; a++ {
		name := names[a]
		fmt.Printf("  %s: %d paragraphs\n", name, alignCounts[a])
	}

	// Check body text paragraphs that should be justified
	fmt.Println("\nBody text paragraphs (should be justify):")
	for i := 136; i <= 150 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			t := r.Text
			if len([]rune(t)) > 40 {
				t = string([]rune(t)[:40]) + "..."
			}
			text += t
		}
		fmt.Printf("  P%d: align=%d heading=%d list=%v text=%q\n",
			i, p.Props.Alignment, p.HeadingLevel, p.IsListItem, text)
	}

	// Check line spacing
	fmt.Println("\nLine spacing samples:")
	for i := 136; i <= 145 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		fmt.Printf("  P%d: lineSpacing=%d lineRule=%d spaceBefore=%d spaceAfter=%d indentFirst=%d\n",
			i, p.Props.LineSpacing, p.Props.LineRule, p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.IndentFirst)
	}
}
