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

	// Check P136 specifically
	if len(fc.Paragraphs) > 136 {
		p := fc.Paragraphs[136]
		fmt.Printf("P136: align=%d alignSet=%v\n", p.Props.Alignment, p.Props.AlignmentSet)
		fmt.Printf("  lineSpacing=%d indentFirst=%d\n", p.Props.LineSpacing, p.Props.IndentFirst)
	}

	// Count alignment distribution after merge
	alignCounts := map[uint8]int{}
	for _, p := range fc.Paragraphs {
		alignCounts[p.Props.Alignment]++
	}
	fmt.Println("\nFinal alignment distribution:")
	names := []string{"left(0)", "center(1)", "right(2)", "justify(3)"}
	for a := uint8(0); a <= 3; a++ {
		fmt.Printf("  %s: %d\n", names[a], alignCounts[a])
	}
}
