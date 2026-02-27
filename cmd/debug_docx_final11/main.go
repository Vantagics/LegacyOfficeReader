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

	// Show the TOC heading and first few TOC entries with full details
	fmt.Println("=== TOC Area (P[96]-P[120]) ===")
	for i := 96; i <= 120 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
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
		if p.Props.AlignmentSet {
			align += "*"
		}
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}
		fmt.Printf("P[%d] %s indent=%d/%d/%d sp=%d/%d line=%d/%d%s: %q\n",
			i, align, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule,
			flags, text)
		for j, r := range p.Runs {
			rText := r.Text
			if len(rText) > 60 {
				rText = rText[:60] + "..."
			}
			fmt.Printf("  Run[%d]: font=%q sz=%d bold=%v color=%q: %q\n",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Color, rText)
		}
	}

	// Show the "修订记录" heading
	fmt.Println("\n=== 修订记录 heading (P[66]) ===")
	if 66 < len(fc.Paragraphs) {
		p := fc.Paragraphs[66]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		fmt.Printf("P[66] H%d align=%d indent=%d/%d/%d sp=%d/%d line=%d/%d: %q\n",
			p.HeadingLevel, p.Props.Alignment, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule, text)
		for j, r := range p.Runs {
			fmt.Printf("  Run[%d]: font=%q sz=%d bold=%v: %q\n",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Text)
		}
	}
}
