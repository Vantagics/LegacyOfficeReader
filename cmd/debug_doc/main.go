package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("LID: 0x%04X, Codepage: %d\n", d.GetLid(), d.GetCodepage())

	fmt.Println("\n=== FONTS ===")
	fonts := d.GetFonts()
	for i, f := range fonts {
		fmt.Printf("  Font %d: %q\n", i, f)
	}

	fmt.Println("\n=== STYLES ===")
	styleNames := d.GetStyles()
	styleSTIs := d.GetStyleSTIs()
	for i := 0; i < len(styleNames); i++ {
		if styleNames[i] != "" {
			sti := uint16(0)
			if i < len(styleSTIs) {
				sti = styleSTIs[i]
			}
			fmt.Printf("  Style %d: %q (sti=%d)\n", i, styleNames[i], sti)
		}
	}

	fmt.Println("\n=== IMAGES ===")
	images := d.GetImages()
	fmt.Printf("Image count: %d\n", len(images))

	fmt.Println("\n=== DEBUG: CharRun/ParaRun ranges ===")
	d.DebugRanges()

	fmt.Println("\n=== FORMATTED CONTENT ===")
	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content available")
		return
	}

	fmt.Printf("Paragraph count: %d\n", len(fc.Paragraphs))
	// Show table cell widths
	for i, p := range fc.Paragraphs {
		if p.TableRowEnd && len(p.CellWidths) > 0 {
			fmt.Printf("  Table row %d: CellWidths=%v\n", i, p.CellWidths)
		}
	}
	shown := 0
	for i := 0; i < len(fc.Paragraphs) && i < 300; i++ {
		p := fc.Paragraphs[i]
		hasText := false
		for _, r := range p.Runs {
			if r.Text != "" {
				hasText = true
				break
			}
		}
		if !hasText && p.HeadingLevel == 0 && !p.InTable && !p.IsListItem && !p.HasPageBreak && !p.PageBreakBefore && !p.IsTOC && !p.IsSectionBreak {
			continue
		}
		shown++
		fmt.Printf("\n--- Para %d ---\n", i)
		fmt.Printf("  H=%d List=%v/%d/%d Table=%v/%v PBB=%v PB=%v TOC=%v/%d SB=%v/%d\n",
			p.HeadingLevel, p.IsListItem, p.ListType, p.ListLevel,
			p.InTable, p.TableRowEnd, p.PageBreakBefore, p.HasPageBreak,
			p.IsTOC, p.TOCLevel, p.IsSectionBreak, p.SectionType)
		fmt.Printf("  Align=%d IndL=%d IndR=%d IndF=%d SpB=%d SpA=%d LS=%d LR=%d\n",
			p.Props.Alignment, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule)
		for j, r := range p.Runs {
			text := r.Text
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			text = strings.ReplaceAll(text, "\x01", "[IMG]")
			fmt.Printf("  R%d: %q F=%q S=%d B=%v I=%v U=%d C=%q\n",
				j, text, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Italic, r.Props.Underline, r.Props.Color)
		}
	}

	fmt.Println("\n=== HEADERS ===")
	// Debug section breaks
	fmt.Println("\n=== SECTION BREAKS ===")
	d.DebugSections()
	sectionBreakCount := 0
	for i, p := range fc.Paragraphs {
		if p.IsSectionBreak {
			sectionBreakCount++
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 40 {
				text = text[:40] + "..."
			}
			fmt.Printf("  Para %d: SectionType=%d Text=%q\n", i, p.SectionType, text)
		}
	}
	fmt.Printf("  Total section breaks: %d\n", sectionBreakCount)
	// Also check for page breaks with 0x0C
	pageBreakCount := 0
	for i, p := range fc.Paragraphs {
		if p.HasPageBreak {
			pageBreakCount++
			if pageBreakCount <= 10 {
				text := ""
				for _, r := range p.Runs {
					text += r.Text
				}
				if len(text) > 40 {
					text = text[:40] + "..."
				}
				fmt.Printf("  PageBreak at para %d: %q\n", i, text)
			}
		}
	}
	fmt.Printf("  Total page breaks: %d\n", pageBreakCount)
	fmt.Println()
	fmt.Println("=== HEADERS ===")
	for i, h := range fc.Headers {
		fmt.Printf("  Header %d: %q\n", i, h)
		fmt.Printf("  Header %d hex: ", i)
		for _, r := range h {
			fmt.Printf("[U+%04X]", r)
		}
		fmt.Println()
	}
	fmt.Println("\n=== FOOTERS ===")
	for i, f := range fc.Footers {
		fmt.Printf("  Footer %d: %q\n", i, f)
		fmt.Printf("  Footer %d hex: ", i)
		for _, r := range f {
			fmt.Printf("[U+%04X]", r)
		}
		fmt.Println()
	}
}
