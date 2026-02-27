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

	// Show paragraphs 0-50 with key info
	for i := 0; i < 50 && i < len(paras); i++ {
		p := paras[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		// Clean for display
		if len(text) > 60 {
			text = text[:60] + "..."
		}
		clean := ""
		for _, c := range text {
			if c == '\t' {
				clean += "\\t"
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
		if p.TextBoxText != "" {
			flags += fmt.Sprintf(" TXBX=%q", p.TextBoxText)
		}
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.InTable {
			flags += " TBL"
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}

		fmt.Printf("P[%3d] align=%d text=%q%s\n", i, p.Props.Alignment, clean, flags)
	}
}
