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

	// Print table paragraphs with IsTableCellEnd
	for i := 67; i <= 95; i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 40 {
			text = text[:40] + "..."
		}
		text = strings.ReplaceAll(text, "\n", "\\n")
		text = strings.ReplaceAll(text, "\r", "\\r")
		
		flags := ""
		if p.TableRowEnd {
			flags += " ROW-END"
		}
		if p.IsTableCellEnd {
			flags += " CELL-END"
		}
		fmt.Printf("P%d: InTable=%v%s text=%q\n", i, p.InTable, flags, text)
	}
}
