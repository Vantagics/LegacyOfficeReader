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

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "DOC: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	fmt.Printf("Headers (%d):\n", len(fc.Headers))
	for i, h := range fc.Headers {
		fmt.Printf("  [%d] %q\n", i, h)
	}
	fmt.Printf("Footers (%d):\n", len(fc.Footers))
	for i, f := range fc.Footers {
		fmt.Printf("  [%d] %q\n", i, f)
	}

	fmt.Printf("\nParagraphs: %d\n", len(fc.Paragraphs))
	// Show first 15 paragraphs to check title page
	for i := 0; i < 15 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		fmt.Printf("  P[%d]: heading=%d table=%v drawn=%v text=%q\n",
			i, p.HeadingLevel, p.InTable, p.DrawnImages, text)
	}
}
