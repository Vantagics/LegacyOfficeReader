package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, _ := os.Open("testfie/test.doc")
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
	fmt.Printf("Headers: %d, Footers: %d\n", len(fc.Headers), len(fc.Footers))
	fmt.Printf("HeadersRaw: %d, FootersRaw: %d\n", len(fc.HeadersRaw), len(fc.FootersRaw))

	// Show ALL paragraphs with their properties
	pageNum := 1
	fmt.Printf("\n=== PAGE %d ===\n", pageNum)
	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}

		flags := ""
		if p.InTable { flags += "T" }
		if p.TableRowEnd { flags += "RE" }
		if p.IsTableCellEnd { flags += "CE" }
		if p.HasPageBreak { flags += "PB" }
		if p.PageBreakBefore { flags += "PBB" }
		if p.IsSectionBreak { flags += fmt.Sprintf("S%d", p.SectionType) }
		if p.HeadingLevel > 0 { flags += fmt.Sprintf("H%d", p.HeadingLevel) }
		if p.IsListItem { flags += "Li" }
		if p.IsTOC { flags += fmt.Sprintf("TOC%d", p.TOCLevel) }
		if len(p.DrawnImages) > 0 { flags += fmt.Sprintf("D%v", p.DrawnImages) }
		if p.TextBoxText != "" { flags += "TX" }

		align := []string{"L", "C", "R", "J"}[p.Props.Alignment]
		if !p.Props.AlignmentSet { align = "-" }

		display := text
		if len(display) > 60 { display = display[:60] + "..." }
		display = strings.ReplaceAll(display, "\t", "→")

		empty := strings.TrimSpace(text) == "" && len(p.DrawnImages) == 0 && p.TextBoxText == ""

		marker := " "
		if empty { marker = "·" }

		fmt.Printf("%sP[%3d] %s %-12s %q\n", marker, i, align, flags, display)

		if p.HasPageBreak || p.IsSectionBreak {
			pageNum++
			fmt.Printf("\n=== PAGE %d ===\n", pageNum)
		}
	}

	// Show footer details
	fmt.Println("\n=== FOOTERS ===")
	for i, f := range fc.Footers {
		fmt.Printf("Footer[%d]: %q\n", i, f)
	}
	for i, f := range fc.FootersRaw {
		fmt.Printf("FooterRaw[%d]: len=%d\n", i, len(f))
	}
	fmt.Println("\n=== HEADERS ===")
	for i, h := range fc.Headers {
		fmt.Printf("Header[%d]: %q\n", i, h)
	}
}
