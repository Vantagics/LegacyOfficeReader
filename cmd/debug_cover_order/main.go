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
	limit := 20
	if limit > len(fc.Paragraphs) {
		limit = len(fc.Paragraphs)
	}
	for i := 0; i < limit; i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		// Truncate text for display
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		fmt.Printf("P[%d] heading=%d secBreak=%v pageBreak=%v drawn=%v textbox=%q align=%d text=%q\n",
			i, p.HeadingLevel, p.IsSectionBreak, p.HasPageBreak, p.DrawnImages, p.TextBoxText, p.Props.Alignment, text)
	}
}
