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
	
	// Show context around inline images and topology diagrams
	targets := []int{138, 139, 140, 141, 142, 143, 144, 145, 155, 156, 157, 158, 159,
		214, 215, 216, 217, 237, 238, 239, 240, 259, 260, 261, 262}
	
	for _, i := range targets {
		if i >= len(fc.Paragraphs) { continue }
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		runes := []rune(text)
		if len(runes) > 60 {
			text = string(runes[:60]) + "..."
		}
		extra := ""
		if p.HeadingLevel > 0 {
			extra += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if len(p.DrawnImages) > 0 {
			extra += fmt.Sprintf(" drawn=%v", p.DrawnImages)
		}
		if p.HasPageBreak {
			extra += " PAGE_BREAK"
		}
		if p.PageBreakBefore {
			extra += " PAGE_BREAK_BEFORE"
		}
		fmt.Printf("P[%d] align=%d%s text=%q\n", i, p.Props.Alignment, extra, text)
	}
}
