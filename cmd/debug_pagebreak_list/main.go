package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "parse: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("no formatted content")
		return
	}

	fmt.Println("Paragraphs with PageBreakBefore=true or HasPageBreak=true:")
	for i, p := range fc.Paragraphs {
		if !p.PageBreakBefore && !p.HasPageBreak {
			continue
		}
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.TrimSpace(text)
		if len(text) > 60 {
			text = text[:60] + "..."
		}
		flags := ""
		if p.PageBreakBefore {
			flags += " PageBreakBefore"
		}
		if p.HasPageBreak {
			flags += " HasPageBreak"
		}
		fmt.Printf("[%3d] %s  text=%q\n", i+1, flags, text)
	}
}
