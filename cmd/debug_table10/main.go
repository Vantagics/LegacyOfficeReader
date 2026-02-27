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

	// Get para runs for debugging
	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Print table paragraphs with more detail
	for i := 67; i <= 95; i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 50 {
			text = text[:50] + "..."
		}
		text = strings.ReplaceAll(text, "\n", "\\n")
		text = strings.ReplaceAll(text, "\r", "\\r")
		
		rowEnd := ""
		if p.TableRowEnd {
			rowEnd = " [ROW-END]"
		}
		cellW := ""
		if len(p.CellWidths) > 0 {
			cellW = fmt.Sprintf(" cw=%v", p.CellWidths)
		}
		fmt.Printf("P%d: InTable=%v%s%s text=%q\n", i, p.InTable, rowEnd, cellW, text)
	}
}
