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

	fmt.Printf("Total paragraphs: %d\n\n", len(fc.Paragraphs))

	for i, p := range fc.Paragraphs {
		text := ""
		hasInlineImg := false
		for _, r := range p.Runs {
			text += r.Text
			if r.ImageRef >= 0 {
				hasInlineImg = true
			}
		}
		text = strings.TrimSpace(text)
		text = strings.ReplaceAll(text, "\r", "")
		text = strings.ReplaceAll(text, "\n", "")

		pType := "TEXT"
		if text == "" && !hasInlineImg && len(p.DrawnImages) == 0 && p.TextBoxText == "" {
			pType = "EMPTY"
		}
		if hasInlineImg || len(p.DrawnImages) > 0 {
			pType = "IMAGE"
		}
		if p.TextBoxText != "" {
			pType = "TEXTBOX"
		}
		if p.HasPageBreak || p.PageBreakBefore {
			pType = "PAGEBRK"
		}
		if p.HeadingLevel > 0 {
			pType = fmt.Sprintf("H%d", p.HeadingLevel)
		}
		if p.IsTOC {
			pType = "TOC"
		}
		if p.IsSectionBreak {
			pType = fmt.Sprintf("SECT(%d)", p.SectionType)
		}
		if p.InTable {
			pType = "TABLE"
		}

		extra := ""
		if len(p.DrawnImages) > 0 {
			extra += fmt.Sprintf(" drawn=%v", p.DrawnImages)
		}
		if p.IsListItem {
			extra += fmt.Sprintf(" list(type=%d,lvl=%d,ilfo=%d)", p.ListType, p.ListLevel, p.ListIlfo)
		}

		display := text
		if len(display) > 60 {
			display = display[:60] + "..."
		}

		fmt.Printf("[%3d] %-10s %s%s\n", i+1, pType, display, extra)
	}
}
