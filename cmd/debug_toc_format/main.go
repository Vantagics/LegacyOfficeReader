package main

import (
	"fmt"
	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}
	
	// Show formatting details for key paragraphs
	targets := []int{33, 40, 66, 96, 97, 98, 135, 152, 155}
	for _, i := range targets {
		if i >= len(fc.Paragraphs) { continue }
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		runes := []rune(text)
		if len(runes) > 50 {
			text = string(runes[:50]) + "..."
		}
		fmt.Printf("\nP[%d] align=%d heading=%d list=%v spacing=%d/%d line=%d/%d indent=%d/%d/%d\n",
			i, p.Props.Alignment, p.HeadingLevel, p.IsListItem,
			p.Props.SpaceBefore, p.Props.SpaceAfter,
			p.Props.LineSpacing, p.Props.LineRule,
			p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst)
		fmt.Printf("  text=%q\n", text)
		for j, r := range p.Runs {
			if j > 3 { break }
			fmt.Printf("  Run[%d]: font=%q size=%d bold=%v italic=%v color=%q text=%q\n",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Italic, r.Props.Color, truncate(r.Text, 30))
		}
	}
}

func truncate(s string, n int) string {
	runes := []rune(s)
	if len(runes) > n {
		return string(runes[:n]) + "..."
	}
	return s
}
