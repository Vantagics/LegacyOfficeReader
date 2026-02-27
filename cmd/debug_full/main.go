package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Println("Error:", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))
	fmt.Printf("Headers: %v\n", fc.Headers)
	fmt.Printf("Footers: %v\n\n", fc.Footers)

	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 70 {
			text = text[:70] + "..."
		}
		text = strings.ReplaceAll(text, "\n", "\\n")
		text = strings.ReplaceAll(text, "\r", "\\r")
		text = strings.ReplaceAll(text, "\t", "\\t")

		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.HasPageBreak {
			flags += " PB"
		}
		if p.PageBreakBefore {
			flags += " PBB"
		}
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" SECT(%d)", p.SectionType)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" IMG=%v", p.DrawnImages)
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}
		if p.IsListItem {
			flags += fmt.Sprintf(" LIST(%d,%d)", p.ListType, p.ListLevel)
		}
		if p.InTable {
			flags += " TBL"
		}
		if p.TableRowEnd {
			flags += " ROWEND"
		}
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				flags += fmt.Sprintf(" INLINE=%d", r.ImageRef)
			}
		}

		align := []string{"L", "C", "R", "J"}[p.Props.Alignment]
		font := ""
		if len(p.Runs) > 0 && p.Runs[0].Props.FontName != "" {
			font = " " + p.Runs[0].Props.FontName
		}
		size := ""
		if len(p.Runs) > 0 && p.Runs[0].Props.FontSize > 0 {
			size = fmt.Sprintf(" %dpt", p.Runs[0].Props.FontSize/2)
		}

		fmt.Printf("P%d [%s%s%s%s]: %q\n", i, align, font, size, flags, text)
	}
}
