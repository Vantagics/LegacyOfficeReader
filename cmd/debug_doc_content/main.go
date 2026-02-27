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

	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))
	fmt.Printf("Images: %d\n", len(d.GetImages()))
	fmt.Printf("Headers: %d, Footers: %d\n", len(fc.Headers), len(fc.Footers))
	fmt.Printf("HeaderEntries: %d, FooterEntries: %d\n", len(fc.HeaderEntries), len(fc.FooterEntries))
	fmt.Println()

	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
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
		if p.TableRowEnd {
			flags += " ROW_END"
		}
		if p.PageBreakBefore {
			flags += " PB_BEFORE"
		}
		if p.HasPageBreak {
			flags += " HAS_PB"
		}
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" SECT(type=%d)", p.SectionType)
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC(lvl=%d)", p.TOCLevel)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN_IMG%v", p.DrawnImages)
		}
		if p.TextBoxText != "" {
			flags += fmt.Sprintf(" TEXTBOX=%q", p.TextBoxText)
		}

		align := "L"
		switch p.Props.Alignment {
		case 1:
			align = "C"
		case 2:
			align = "R"
		case 3:
			align = "J"
		}
		if p.Props.AlignmentSet {
			align += "*"
		}

		spacing := fmt.Sprintf("before=%d after=%d line=%d/%d", p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule)
		indent := fmt.Sprintf("L=%d R=%d F=%d", p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst)

		// Truncate text for display
		display := text
		if len(display) > 80 {
			display = display[:80] + "..."
		}

		fmt.Printf("P[%d] align=%s %s %s%s\n", i, align, spacing, indent, flags)
		
		// Show runs with formatting
		for j, r := range p.Runs {
			runText := r.Text
			if len(runText) > 60 {
				runText = runText[:60] + "..."
			}
			runFlags := ""
			if r.Props.Bold {
				runFlags += " B"
			}
			if r.Props.Italic {
				runFlags += " I"
			}
			if r.Props.Underline > 0 {
				runFlags += fmt.Sprintf(" U%d", r.Props.Underline)
			}
			if r.Props.Color != "" {
				runFlags += " #" + r.Props.Color
			}
			if r.ImageRef >= 0 {
				runFlags += fmt.Sprintf(" IMG[%d]", r.ImageRef)
			}
			fmt.Printf("  R[%d] font=%q sz=%d%s: %q\n", j, r.Props.FontName, r.Props.FontSize, runFlags, runText)
		}
		fmt.Println()
	}

	// Print header/footer info
	for i, he := range fc.HeaderEntries {
		fmt.Printf("Header[%d] type=%s text=%q images=%v\n", i, he.Type, he.Text, he.Images)
		fmt.Printf("  raw=%q\n", he.RawText)
	}
	for i, fe := range fc.FooterEntries {
		fmt.Printf("Footer[%d] type=%s text=%q images=%v\n", i, fe.Type, fe.Text, fe.Images)
		fmt.Printf("  raw=%q\n", fe.RawText)
	}

	// Print section info
	fmt.Println("\n--- Section breaks ---")
	for i, p := range fc.Paragraphs {
		if p.IsSectionBreak {
			fmt.Printf("P[%d] SectionBreak type=%d\n", i, p.SectionType)
		}
	}
}
