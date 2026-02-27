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

	fmt.Printf("Total paragraphs: %d\n\n", len(fc.Paragraphs))

	// Show first 20 paragraphs in detail to understand title page structure
	for i := 0; i < len(fc.Paragraphs) && i < 25; i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}

		flags := ""
		if p.InTable { flags += " TABLE" }
		if p.TableRowEnd { flags += " ROWEND" }
		if p.IsTableCellEnd { flags += " CELLEND" }
		if p.HasPageBreak { flags += " PAGEBREAK" }
		if p.PageBreakBefore { flags += " PBB" }
		if p.IsSectionBreak { flags += fmt.Sprintf(" SECT(%d)", p.SectionType) }
		if p.HeadingLevel > 0 { flags += fmt.Sprintf(" H%d", p.HeadingLevel) }
		if p.IsListItem { flags += fmt.Sprintf(" LIST(%d,L%d)", p.ListType, p.ListLevel) }
		if p.IsTOC { flags += fmt.Sprintf(" TOC%d", p.TOCLevel) }
		if len(p.DrawnImages) > 0 { flags += fmt.Sprintf(" DRAWN=%v", p.DrawnImages) }
		if p.TextBoxText != "" { flags += fmt.Sprintf(" TXBX=%q", p.TextBoxText) }

		align := []string{"L", "C", "R", "J"}[p.Props.Alignment]
		if !p.Props.AlignmentSet { align = "-" }

		// Show run details
		runInfo := ""
		for j, r := range p.Runs {
			ri := fmt.Sprintf("R%d[", j)
			if r.Props.Bold { ri += "B" }
			if r.Props.Italic { ri += "I" }
			if r.Props.FontSize > 0 { ri += fmt.Sprintf("sz%d", r.Props.FontSize) }
			if r.Props.FontName != "" { ri += fmt.Sprintf("f=%s", r.Props.FontName) }
			if r.ImageRef >= 0 { ri += fmt.Sprintf("img%d", r.ImageRef) }
			ri += "]"
			if j < 3 { runInfo += ri + " " }
		}

		if len(text) > 50 {
			text = text[:50] + "..."
		}
		fmt.Printf("P[%3d] align=%s%s %s text=%q\n", i, align, flags, runInfo, text)
	}

	// Show page break locations
	fmt.Println("\n--- Page breaks ---")
	for i, p := range fc.Paragraphs {
		if p.HasPageBreak || p.PageBreakBefore {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 40 { text = text[:40] + "..." }
			fmt.Printf("P[%d]: HasPageBreak=%v PBB=%v text=%q\n", i, p.HasPageBreak, p.PageBreakBefore, text)
		}
	}

	// Show section breaks
	fmt.Println("\n--- Section breaks ---")
	for i, p := range fc.Paragraphs {
		if p.IsSectionBreak {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 40 { text = text[:40] + "..." }
			fmt.Printf("P[%d]: SectionType=%d text=%q\n", i, p.SectionType, text)
		}
	}

	// Show headings
	fmt.Println("\n--- Headings ---")
	for i, p := range fc.Paragraphs {
		if p.HeadingLevel > 0 {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 { text = text[:60] + "..." }
			fmt.Printf("P[%d]: H%d text=%q\n", i, p.HeadingLevel, text)
		}
	}

	// Headers/footers
	fmt.Println("\n--- Headers ---")
	for i, h := range fc.Headers {
		fmt.Printf("  Header[%d]: %q\n", i, h)
	}
	fmt.Println("--- Footers ---")
	for i, f := range fc.Footers {
		fmt.Printf("  Footer[%d]: %q\n", i, f)
	}
	fmt.Println("--- FootersRaw ---")
	for i, f := range fc.FootersRaw {
		raw := ""
		for _, r := range f {
			if r < 0x20 && r != '\t' {
				raw += fmt.Sprintf("[%02X]", r)
			} else {
				raw += string(r)
			}
		}
		fmt.Printf("  FooterRaw[%d]: %s\n", i, raw)
	}

	// Count empty paragraphs at start
	fmt.Println("\n--- Empty paragraphs at start ---")
	for i := 0; i < 15 && i < len(fc.Paragraphs); i++ {
		text := ""
		for _, r := range fc.Paragraphs[i].Runs {
			text += r.Text
		}
		isEmpty := strings.TrimSpace(text) == ""
		fmt.Printf("P[%d]: empty=%v len=%d\n", i, isEmpty, len(text))
	}
}
