package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	document, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	fc := document.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Show all list items with their list type and level
	fmt.Println("=== All List Items ===")
	for i, p := range fc.Paragraphs {
		if p.IsListItem {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			heading := ""
			if p.HeadingLevel > 0 {
				heading = fmt.Sprintf(" H%d", p.HeadingLevel)
			}
			fmt.Printf("P[%d] ListType=%d ListLevel=%d%s: %q\n",
				i, p.ListType, p.ListLevel, heading, text)
		}
	}
}
