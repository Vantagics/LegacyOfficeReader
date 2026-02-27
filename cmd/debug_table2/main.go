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

	// Show all table paragraphs with full detail
	fmt.Printf("=== Table Paragraphs Detail ===\n")
	for i, p := range fc.Paragraphs {
		if !p.InTable {
			continue
		}
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 40 {
			text = text[:40] + "..."
		}
		fmt.Printf("P%d: %q inTable=%v rowEnd=%v cellWidths=%v align=%d\n",
			i, text, p.InTable, p.TableRowEnd, p.CellWidths, p.Props.Alignment)
	}
}
