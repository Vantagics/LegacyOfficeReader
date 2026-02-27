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
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))
	fmt.Printf("Headers: %d\n", len(fc.Headers))
	fmt.Printf("Footers: %d\n", len(fc.Footers))

	// Count by type
	headings := map[uint8]int{}
	listItems := 0
	tableParas := 0
	tocParas := 0
	pageBreaks := 0
	sectionBreaks := 0
	drawnImgs := 0
	inlineImgs := 0
	textBoxes := 0
	emptyParas := 0

	for _, p := range fc.Paragraphs {
		if p.HeadingLevel > 0 {
			headings[p.HeadingLevel]++
		}
		if p.IsListItem {
			listItems++
		}
		if p.InTable {
			tableParas++
		}
		if p.IsTOC {
			tocParas++
		}
		if p.HasPageBreak {
			pageBreaks++
		}
		if p.IsSectionBreak {
			sectionBreaks++
		}
		if len(p.DrawnImages) > 0 {
			drawnImgs++
		}
		if p.TextBoxText != "" {
			textBoxes++
		}
		hasInline := false
		isEmpty := true
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				hasInline = true
			}
			text := r.Text
			for _, c := range text {
				if c > 0x20 && c != 0x01 && c != 0x07 && c != 0x08 && c != 0x0D {
					isEmpty = false
				}
			}
		}
		if hasInline {
			inlineImgs++
		}
		if isEmpty && p.TextBoxText == "" && len(p.DrawnImages) == 0 {
			emptyParas++
		}
	}

	fmt.Printf("\nHeadings:\n")
	for lvl := uint8(1); lvl <= 9; lvl++ {
		if c, ok := headings[lvl]; ok {
			fmt.Printf("  H%d: %d\n", lvl, c)
		}
	}
	fmt.Printf("List items: %d\n", listItems)
	fmt.Printf("Table paragraphs: %d\n", tableParas)
	fmt.Printf("TOC paragraphs: %d\n", tocParas)
	fmt.Printf("Page breaks: %d\n", pageBreaks)
	fmt.Printf("Section breaks: %d\n", sectionBreaks)
	fmt.Printf("Paragraphs with drawn images: %d\n", drawnImgs)
	fmt.Printf("Paragraphs with inline images: %d\n", inlineImgs)
	fmt.Printf("Text boxes: %d\n", textBoxes)
	fmt.Printf("Empty paragraphs: %d\n", emptyParas)

	// Show first 20 paragraphs
	fmt.Println("\n=== First 20 paragraphs ===")
	for i := 0; i < 20 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.IsListItem {
			flags += fmt.Sprintf(" LIST(type=%d,lvl=%d)", p.ListType, p.ListLevel)
		}
		if p.InTable {
			flags += " TABLE"
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}
		if p.HasPageBreak {
			flags += " PAGEBRK"
		}
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" SECBRK(%d)", p.SectionType)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN%v", p.DrawnImages)
		}
		if p.TextBoxText != "" {
			flags += fmt.Sprintf(" TXBX(%s)", p.TextBoxText[:min(20, len(p.TextBoxText))])
		}
		align := []string{"left", "center", "right", "both"}[p.Props.Alignment]
		fmt.Printf("P[%d] align=%s%s: %q\n", i, align, flags, text)
	}

	// Show heading paragraphs
	fmt.Println("\n=== All Headings ===")
	for i, p := range fc.Paragraphs {
		if p.HeadingLevel > 0 {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 80 {
				text = text[:80] + "..."
			}
			listInfo := ""
			if p.IsListItem {
				listInfo = fmt.Sprintf(" LIST(type=%d,lvl=%d)", p.ListType, p.ListLevel)
			}
			fmt.Printf("P[%d] H%d%s: %q\n", i, p.HeadingLevel, listInfo, text)
		}
	}

	// Show footer info
	fmt.Println("\n=== Footers ===")
	for i, f := range fc.Footers {
		fmt.Printf("Footer[%d]: %q\n", i, f)
	}
	for i, f := range fc.FootersRaw {
		fmt.Printf("FooterRaw[%d]: %q\n", i, f)
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
