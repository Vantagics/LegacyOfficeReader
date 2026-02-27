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

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	fmt.Printf("=== HEADERS (%d) ===\n", len(fc.Headers))
	for i, h := range fc.Headers {
		fmt.Printf("  Header[%d]: %q\n", i, h)
	}
	fmt.Printf("=== FOOTERS (%d) ===\n", len(fc.Footers))
	for i, f := range fc.Footers {
		fmt.Printf("  Footer[%d]: %q\n", i, f)
	}
	fmt.Printf("=== HEADERS RAW (%d) ===\n", len(fc.HeadersRaw))
	for i, h := range fc.HeadersRaw {
		fmt.Printf("  HeaderRaw[%d]: %q\n", i, h)
	}
	fmt.Printf("=== FOOTERS RAW (%d) ===\n", len(fc.FootersRaw))
	for i, f := range fc.FootersRaw {
		fmt.Printf("  FooterRaw[%d]: %q\n", i, f)
	}

	fmt.Printf("\n=== PARAGRAPHS (%d) ===\n", len(fc.Paragraphs))
	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.TrimRight(text, "\r\n")
		
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.InTable {
			flags += " TABLE"
		}
		if p.TableRowEnd {
			flags += " ROWEND"
		}
		if p.IsTableCellEnd {
			flags += " CELLEND"
		}
		if p.PageBreakBefore {
			flags += " PGBRK-BEFORE"
		}
		if p.HasPageBreak {
			flags += " HAS-PGBRK"
		}
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" SECBRK(type=%d)", p.SectionType)
		}
		if p.IsListItem {
			flags += fmt.Sprintf(" LIST(type=%d,lvl=%d)", p.ListType, p.ListLevel)
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC(lvl=%d)", p.TOCLevel)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN-IMG%v", p.DrawnImages)
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
		if !p.Props.AlignmentSet {
			align = "?" + align
		}
		
		spacing := fmt.Sprintf("before=%d after=%d line=%d/%d", p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule)
		indent := fmt.Sprintf("L=%d R=%d F=%d", p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst)
		
		preview := text
		if len(preview) > 80 {
			preview = preview[:80] + "..."
		}
		
		fmt.Printf("P[%d] align=%s %s %s%s\n", i, align, spacing, indent, flags)
		if len(p.Runs) > 0 {
			for j, r := range p.Runs {
				rtext := strings.TrimRight(r.Text, "\r\n")
				if len(rtext) > 60 {
					rtext = rtext[:60] + "..."
				}
				rfmt := ""
				if r.Props.Bold {
					rfmt += "B"
				}
				if r.Props.Italic {
					rfmt += "I"
				}
				if r.Props.Underline > 0 {
					rfmt += "U"
				}
				if r.Props.FontSize > 0 {
					rfmt += fmt.Sprintf(" sz=%d", r.Props.FontSize)
				}
				if r.Props.FontName != "" {
					rfmt += fmt.Sprintf(" fn=%q", r.Props.FontName)
				}
				if r.Props.Color != "" {
					rfmt += fmt.Sprintf(" clr=%s", r.Props.Color)
				}
				if r.ImageRef >= 0 {
					rfmt += fmt.Sprintf(" IMG=%d", r.ImageRef)
				}
				fmt.Printf("  R[%d] %s %q\n", j, rfmt, rtext)
			}
		}
	}

	fmt.Printf("\n=== IMAGES (%d) ===\n", len(d.GetImages()))
	for i, img := range d.GetImages() {
		fmt.Printf("  Image[%d]: ext=%q size=%d\n", i, img.Extension(), len(img.Data))
	}
}
