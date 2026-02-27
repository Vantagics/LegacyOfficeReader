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

	images := d.GetImages()
	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))
	fmt.Printf("Total images: %d\n", len(images))
	fmt.Printf("Headers: %v\n", fc.Headers)
	fmt.Printf("Footers: %v\n", fc.Footers)
	fmt.Println()

	// Print all paragraphs with their properties
	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		// Clean control chars for display
		display := strings.Map(func(r rune) rune {
			if r < 0x20 && r != '\t' {
				return '·'
			}
			return r
		}, text)
		if len([]rune(display)) > 80 {
			display = string([]rune(display)[:80]) + "..."
		}

		flags := ""
		if p.InTable {
			flags += " [TABLE"
			if p.TableRowEnd {
				flags += " ROW-END"
			}
			if p.IsTableCellEnd {
				flags += " CELL-END"
			}
			if len(p.CellWidths) > 0 {
				flags += fmt.Sprintf(" widths=%v", p.CellWidths)
			}
			flags += "]"
		}
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" [H%d]", p.HeadingLevel)
		}
		if p.IsListItem {
			flags += fmt.Sprintf(" [LIST type=%d lvl=%d]", p.ListType, p.ListLevel)
		}
		if p.HasPageBreak {
			flags += " [PAGEBREAK]"
		}
		if p.PageBreakBefore {
			flags += " [PBB]"
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" [TOC%d]", p.TOCLevel)
		}
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" [SECT type=%d]", p.SectionType)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" [DRAWN=%v]", p.DrawnImages)
		}
		if p.Props.Alignment > 0 {
			aligns := []string{"left", "center", "right", "justify"}
			a := "?"
			if int(p.Props.Alignment) < len(aligns) {
				a = aligns[p.Props.Alignment]
			}
			flags += fmt.Sprintf(" [align=%s]", a)
		}

		// Show run details for first few runs
		runInfo := ""
		for j, r := range p.Runs {
			if j >= 3 {
				runInfo += fmt.Sprintf(" +%d more", len(p.Runs)-3)
				break
			}
			ri := ""
			if r.Props.FontName != "" {
				ri += " font=" + r.Props.FontName
			}
			if r.Props.FontSize > 0 {
				ri += fmt.Sprintf(" sz=%d", r.Props.FontSize)
			}
			if r.Props.Bold {
				ri += " B"
			}
			if r.ImageRef >= 0 {
				ri += fmt.Sprintf(" IMG=%d", r.ImageRef)
			}
			if ri != "" {
				runInfo += fmt.Sprintf(" R%d:{%s}", j, ri)
			}
		}

		fmt.Printf("P%d: %q%s%s\n", i, display, flags, runInfo)
	}
}
