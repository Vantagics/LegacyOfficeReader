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

	// Show paragraphs with drawn images
	fmt.Println("=== Paragraphs with DrawnImages ===")
	for i, p := range paras {
		if len(p.DrawnImages) > 0 {
			text := ""
			for _, r := range p.Runs {
				t := r.Text
				clean := ""
				for _, c := range t {
					if c == '\t' {
						clean += "\\t"
					} else if c < 0x20 {
						clean += fmt.Sprintf("\\x%02x", c)
					} else {
						clean += string(c)
					}
				}
				text += clean
			}
			fmt.Printf("P[%d]: DrawnImages=%v, TextBox=%q, Text=%q\n", i, p.DrawnImages, p.TextBoxText, text)
			fmt.Printf("  Align=%d, Heading=%d, IsSectionBreak=%v\n", p.Props.Alignment, p.HeadingLevel, p.IsSectionBreak)
		}
	}

	// Show paragraphs with textboxes
	fmt.Println("\n=== Paragraphs with TextBoxText ===")
	for i, p := range paras {
		if p.TextBoxText != "" {
			fmt.Printf("P[%d]: TextBox=%q, DrawnImages=%v\n", i, p.TextBoxText, p.DrawnImages)
		}
	}

	// Show paragraphs with inline images (0x01)
	fmt.Println("\n=== Paragraphs with Inline Images ===")
	for i, p := range paras {
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				fmt.Printf("P[%d]: ImageRef=%d, Text=%q\n", i, r.ImageRef, r.Text)
			}
		}
	}

	// Show page breaks and section breaks
	fmt.Println("\n=== Page/Section Breaks ===")
	for i, p := range paras {
		if p.HasPageBreak || p.PageBreakBefore || p.IsSectionBreak {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 50 {
				text = text[:50] + "..."
			}
			fmt.Printf("P[%d]: PageBreak=%v, PageBreakBefore=%v, SectionBreak=%v(type=%d), Text=%q\n",
				i, p.HasPageBreak, p.PageBreakBefore, p.IsSectionBreak, p.SectionType, text)
		}
	}
}
