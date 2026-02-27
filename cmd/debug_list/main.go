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
	paras := fc.Paragraphs

	// Show all paragraphs with list properties
	fmt.Println("=== List Paragraphs ===")
	for i, p := range paras {
		if p.IsListItem {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 80 {
				text = text[:80] + "..."
			}
			fmt.Printf("P[%3d] ListType=%d Level=%d Heading=%d text=%q\n",
				i, p.ListType, p.ListLevel, p.HeadingLevel, text)
		}
	}

	// Also show paragraphs around the area in the screenshot (P[255]-P[278])
	fmt.Println("\n=== P[250]-P[278] Detail ===")
	for i := 250; i < len(paras) && i <= 278; i++ {
		p := paras[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		flags := ""
		if p.IsListItem {
			flags += fmt.Sprintf(" LIST(type=%d,lvl=%d)", p.ListType, p.ListLevel)
		}
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN=%v", p.DrawnImages)
		}
		fmt.Printf("P[%3d] align=%d iLeft=%d iFirst=%d%s text=%q\n",
			i, p.Props.Alignment, p.Props.IndentLeft, p.Props.IndentFirst, flags, text)
	}

	// Show paragraphs around P[195]-P[215] (产品价值 section with sub-items)
	fmt.Println("\n=== P[195]-P[215] Detail ===")
	for i := 195; i < len(paras) && i <= 215; i++ {
		p := paras[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		flags := ""
		if p.IsListItem {
			flags += fmt.Sprintf(" LIST(type=%d,lvl=%d)", p.ListType, p.ListLevel)
		}
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		fmt.Printf("P[%3d] align=%d iLeft=%d iFirst=%d%s text=%q\n",
			i, p.Props.Alignment, p.Props.IndentLeft, p.Props.IndentFirst, flags, text)
	}
}
