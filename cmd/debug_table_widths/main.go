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
		return
	}

	for i, p := range fc.Paragraphs {
		if p.TableRowEnd {
			fmt.Printf("P[%d] ROWEND CellWidths=%v\n", i, p.CellWidths)
		}
	}
}
