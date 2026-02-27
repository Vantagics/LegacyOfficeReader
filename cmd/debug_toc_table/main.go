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
	
	// Show paragraphs 10-70 (TOC and pre-table area)
	for i := 10; i < 100 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		runes := []rune(text)
		if len(runes) > 70 {
			text = string(runes[:70]) + "..."
		}
		extra := ""
		if p.IsTOC {
			extra += fmt.Sprintf(" TOC(%d)", p.TOCLevel)
		}
		if p.HeadingLevel > 0 {
			extra += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.InTable {
			extra += " TABLE"
			if p.TableRowEnd {
				extra += "(ROW_END)"
				extra += fmt.Sprintf(" widths=%v", p.CellWidths)
			}
			if p.IsTableCellEnd {
				extra += "(CELL_END)"
			}
		}
		if p.HasPageBreak {
			extra += " PAGE_BREAK"
		}
		if p.IsSectionBreak {
			extra += fmt.Sprintf(" SECTION(%d)", p.SectionType)
		}
		if p.IsListItem {
			extra += fmt.Sprintf(" LIST(type=%d,lvl=%d)", p.ListType, p.ListLevel)
		}
		fmt.Printf("P[%d] align=%d%s text=%q\n", i, p.Props.Alignment, extra, text)
	}
}
