package main

import (
	"fmt"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		return
	}

	styles := d.GetStyles()
	stis := d.GetStyleSTIs()
	fmt.Printf("Styles (%d):\n", len(styles))
	for i, name := range styles {
		if name != "" || (i < len(stis) && stis[i] != 0) {
			sti := uint16(0)
			if i < len(stis) {
				sti = stis[i]
			}
			fmt.Printf("  [%d] sti=%d name=%q\n", i, sti, name)
		}
	}

	// Check what style the body text paragraphs use
	fc := d.GetFormattedContent()
	fmt.Printf("\nParagraphs 135-155 (body text area):\n")
	for i := 135; i < 155 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 50 {
			text = text[:50] + "..."
		}
		fmt.Printf("  P%d: align=%d spacing=%d lineSpacing=%d heading=%d %q\n",
			i, p.Props.Alignment, p.Props.SpaceAfter, p.Props.LineSpacing, p.HeadingLevel, text)
	}
}
