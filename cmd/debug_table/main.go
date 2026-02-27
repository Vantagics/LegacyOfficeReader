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

	fc := d.GetFormattedContent()

	// Show table paragraphs
	fmt.Printf("=== Table Paragraphs ===\n")
	rowNum := 0
	cellNum := 0
	for i, p := range fc.Paragraphs {
		if !p.InTable {
			if cellNum > 0 {
				fmt.Printf("  (end of table region)\n")
				cellNum = 0
				rowNum = 0
			}
			continue
		}
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 40 {
			text = text[:40] + "..."
		}
		flags := ""
		if p.TableRowEnd {
			flags += " ROWEND"
			if len(p.CellWidths) > 0 {
				flags += fmt.Sprintf(" widths=%v", p.CellWidths)
			}
		}
		fmt.Printf("  P%d: row=%d cell=%d %q%s\n", i, rowNum, cellNum, text, flags)
		if p.TableRowEnd {
			rowNum++
			cellNum = 0
		} else {
			cellNum++
		}
	}
}
