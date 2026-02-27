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

	// Show paragraphs around page breaks and drawn images
	// Focus on P[140]-P[170] (inline images area) and P[210]-P[270] (topology diagrams)
	ranges := [][2]int{{135, 165}, {190, 220}, {235, 270}}
	for _, rng := range ranges {
		fmt.Printf("\n=== P[%d]-P[%d] ===\n", rng[0], rng[1])
		for i := rng[0]; i <= rng[1] && i < len(paras); i++ {
			p := paras[i]
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 80 {
				text = text[:80] + "..."
			}
			clean := ""
			for _, c := range text {
				if c == '\t' {
					clean += "\\t"
				} else if c == '\x01' {
					clean += "\\x01"
				} else if c < 0x20 {
					clean += fmt.Sprintf("\\x%02x", c)
				} else {
					clean += string(c)
				}
			}

			flags := ""
			if p.IsSectionBreak {
				flags += fmt.Sprintf(" SEC(%d)", p.SectionType)
			}
			if p.HasPageBreak {
				flags += " PB"
			}
			if p.PageBreakBefore {
				flags += " PBB"
			}
			if len(p.DrawnImages) > 0 {
				flags += fmt.Sprintf(" DRAWN=%v", p.DrawnImages)
			}
			if p.HeadingLevel > 0 {
				flags += fmt.Sprintf(" H%d", p.HeadingLevel)
			}
			if p.IsTOC {
				flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
			}
			for _, r := range p.Runs {
				if r.ImageRef >= 0 {
					flags += fmt.Sprintf(" IMG=%d", r.ImageRef)
				}
			}

			fmt.Printf("P[%3d] align=%d text=%q%s\n", i, p.Props.Alignment, clean, flags)
		}
	}
}
