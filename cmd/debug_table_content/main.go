package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	fmt.Println("=== Table paragraphs ===")
	for i, p := range fc.Paragraphs {
		if !p.InTable {
			continue
		}
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		flags := ""
		if p.TableRowEnd {
			flags += " ROW_END"
		}
		if p.IsTableCellEnd {
			flags += " CELL_END"
		}
		if len(p.CellWidths) > 0 {
			flags += fmt.Sprintf(" CW=%v", p.CellWidths)
		}
		fmt.Printf("P[%d]%s: %q\n", i, flags, text)
	}
}
