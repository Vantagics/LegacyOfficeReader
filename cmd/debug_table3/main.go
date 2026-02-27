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

	fmt.Printf("Total paragraphs: %d\n\n", len(fc.Paragraphs))

	// Find and dump all table paragraphs
	inTable := false
	tableStart := 0
	tableNum := 0
	for i, p := range fc.Paragraphs {
		if p.InTable && !inTable {
			inTable = true
			tableStart = i
			tableNum++
			fmt.Printf("=== TABLE %d starts at P%d ===\n", tableNum, i)
		}
		if !p.InTable && inTable {
			inTable = false
			fmt.Printf("=== TABLE %d ends at P%d ===\n\n", tableNum, i-1)
		}
		if p.InTable {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			text = strings.ReplaceAll(text, "\n", "\\n")
			text = strings.ReplaceAll(text, "\r", "\\r")
			rowEnd := ""
			if p.TableRowEnd {
				rowEnd = " [ROW-END]"
			}
			cellW := ""
			if len(p.CellWidths) > 0 {
				cellW = fmt.Sprintf(" cellWidths=%v", p.CellWidths)
			}
			fmt.Printf("  P%d: InTable=%v%s%s text=%q\n", i, p.InTable, rowEnd, cellW, text)
		}
	}
	if inTable {
		fmt.Printf("=== TABLE %d ends at P%d (end of doc) ===\n", tableNum, len(fc.Paragraphs)-1)
	}

	_ = tableStart
	fmt.Printf("\n--- Headers: %v\n", fc.Headers)
	fmt.Printf("--- Footers: %v\n", fc.Footers)

	// Also dump first 20 paragraphs to see title page
	fmt.Println("\n=== FIRST 20 PARAGRAPHS ===")
	for i := 0; i < 20 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.HasPageBreak {
			flags += " PAGEBREAK"
		}
		if p.PageBreakBefore {
			flags += " PBB"
		}
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" SECT(%d)", p.SectionType)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DrawnImg=%v", p.DrawnImages)
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}
		if p.IsListItem {
			flags += fmt.Sprintf(" LIST(type=%d,lvl=%d)", p.ListType, p.ListLevel)
		}
		align := []string{"left", "center", "right", "both"}[p.Props.Alignment]
		fmt.Printf("  P%d: align=%s%s text=%q\n", i, align, flags, text)
	}
}
