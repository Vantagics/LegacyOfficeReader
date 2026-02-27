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

	// Count alignment distribution
	alignCounts := map[uint8]int{}
	alignSetCounts := map[uint8]int{}
	for _, p := range fc.Paragraphs {
		alignCounts[p.Props.Alignment]++
		if p.Props.AlignmentSet {
			alignSetCounts[p.Props.Alignment]++
		}
	}

	alignNames := map[uint8]string{0: "left", 1: "center", 2: "right", 3: "both"}
	fmt.Println("Alignment distribution:")
	for a := uint8(0); a <= 3; a++ {
		fmt.Printf("  %s: %d total, %d explicitly set\n", alignNames[a], alignCounts[a], alignSetCounts[a])
	}

	// Show body paragraphs with left alignment that are explicitly set
	fmt.Println("\nExplicitly left-aligned non-heading body paragraphs:")
	count := 0
	for i, p := range fc.Paragraphs {
		if p.Props.Alignment == 0 && p.Props.AlignmentSet && p.HeadingLevel == 0 && !p.InTable && !p.IsTOC {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			fmt.Printf("  P[%d]: %q\n", i, text)
			count++
			if count > 15 {
				fmt.Printf("  ... and more\n")
				break
			}
		}
	}
}
