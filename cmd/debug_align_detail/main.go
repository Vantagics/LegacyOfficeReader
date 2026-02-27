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

	// Check AlignmentSet for body text paragraphs
	fmt.Println("Body text paragraphs alignment details:")
	for i := 136; i <= 150 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			t := r.Text
			if len([]rune(t)) > 30 {
				t = string([]rune(t)[:30]) + "..."
			}
			text += t
		}
		fmt.Printf("  P%d: align=%d alignSet=%v text=%q\n",
			i, p.Props.Alignment, p.Props.AlignmentSet, text)
	}

	// Count AlignmentSet distribution
	setCount := 0
	unsetCount := 0
	for _, p := range fc.Paragraphs {
		if p.Props.AlignmentSet {
			setCount++
		} else {
			unsetCount++
		}
	}
	fmt.Printf("\nAlignmentSet: true=%d false=%d\n", setCount, unsetCount)

	// Check alignment distribution for AlignmentSet=true
	fmt.Println("\nAlignment distribution (AlignmentSet=true only):")
	alignCounts := make(map[uint8]int)
	for _, p := range fc.Paragraphs {
		if p.Props.AlignmentSet {
			alignCounts[p.Props.Alignment]++
		}
	}
	names := []string{"left(0)", "center(1)", "right(2)", "justify(3)"}
	for a := uint8(0); a <= 3; a++ {
		fmt.Printf("  %s: %d\n", names[a], alignCounts[a])
	}
}
