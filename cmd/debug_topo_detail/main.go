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
	
	// Show paragraphs around topology diagrams in detail
	ranges := [][2]int{{215, 230}, {237, 253}, {259, 274}}
	for _, rng := range ranges {
		fmt.Printf("\n--- Range P[%d]-P[%d] ---\n", rng[0], rng[1])
		for i := rng[0]; i <= rng[1] && i < len(fc.Paragraphs); i++ {
			p := fc.Paragraphs[i]
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			runes := []rune(text)
			if len(runes) > 70 {
				text = string(runes[:70]) + "..."
			}
			fmt.Printf("P[%d] align=%d drawn=%v list=%v(type=%d,lvl=%d) heading=%d text=%q\n",
				i, p.Props.Alignment, p.DrawnImages, p.IsListItem, p.ListType, p.ListLevel, p.HeadingLevel, text)
		}
	}
}
