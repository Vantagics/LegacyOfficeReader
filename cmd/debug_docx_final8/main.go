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

	document, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	fc := document.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Show all heading paragraphs with full run details
	fmt.Println("=== All Headings with Run Details ===")
	for i, p := range fc.Paragraphs {
		if p.HeadingLevel > 0 {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			align := "L"
			switch p.Props.Alignment {
			case 1:
				align = "C"
			case 2:
				align = "R"
			case 3:
				align = "J"
			}
			fmt.Printf("\nP[%d] H%d align=%s indent=%d/%d/%d sp=%d/%d line=%d/%d list=%v: %q\n",
				i, p.HeadingLevel, align, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
				p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule,
				p.IsListItem, text)
			for j, r := range p.Runs {
				fmt.Printf("  Run[%d]: font=%q sz=%d bold=%v color=%q\n",
					j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Color)
			}
		}
	}

	// Show style definitions
	fmt.Println("\n=== Style Definitions ===")
	document.DebugStyleProps()
}
