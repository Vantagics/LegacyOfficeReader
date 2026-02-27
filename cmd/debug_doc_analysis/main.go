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

	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))
	fmt.Printf("Headers: %d\n", len(fc.Headers))
	fmt.Printf("Footers: %d\n", len(fc.Footers))
	fmt.Printf("HeadersRaw: %d\n", len(fc.HeadersRaw))
	fmt.Printf("FootersRaw: %d\n", len(fc.FootersRaw))

	for i, h := range fc.Headers {
		fmt.Printf("  Header[%d]: %q\n", i, h)
	}
	for i, f := range fc.Footers {
		fmt.Printf("  Footer[%d]: %q\n", i, f)
	}
	for i, r := range fc.FootersRaw {
		fmt.Printf("  FooterRaw[%d]: %q\n", i, r)
	}

	fmt.Println("\n=== Paragraph Analysis ===")
	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		// Clean for display
		displayText := strings.ReplaceAll(text, "\t", "\\t")
		displayText = strings.ReplaceAll(displayText, "\x01", "[IMG]")
		displayText = strings.ReplaceAll(displayText, "\x08", "[DRAWN]")
		if len(displayText) > 100 {
			displayText = displayText[:100] + "..."
		}

		flags := ""
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" [SECT:%d]", p.SectionType)
		}
		if p.HasPageBreak {
			flags += " [PB]"
		}
		if p.PageBreakBefore {
			flags += " [PBB]"
		}
		if p.InTable {
			flags += " [TBL]"
		}
		if p.TableRowEnd {
			flags += " [ROWEND]"
		}
		if p.IsTableCellEnd {
			flags += " [CELLEND]"
		}
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" [H%d]", p.HeadingLevel)
		}
		if p.IsListItem {
			flags += fmt.Sprintf(" [LIST:%d.%d]", p.ListType, p.ListLevel)
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" [TOC%d]", p.TOCLevel)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" [DRAWN:%v]", p.DrawnImages)
		}
		if p.TextBoxText != "" {
			tbText := p.TextBoxText
			if len(tbText) > 40 {
				tbText = tbText[:40] + "..."
			}
			flags += fmt.Sprintf(" [TXBX:%q]", tbText)
		}
		if p.Props.Alignment == 1 {
			flags += " [CENTER]"
		} else if p.Props.Alignment == 2 {
			flags += " [RIGHT]"
		}

		// Font info from first run
		fontInfo := ""
		if len(p.Runs) > 0 {
			r := p.Runs[0]
			if r.Props.FontName != "" || r.Props.FontSize > 0 {
				fontInfo = fmt.Sprintf(" font=%s/%dhp", r.Props.FontName, r.Props.FontSize)
			}
			if r.Props.Bold {
				fontInfo += " B"
			}
		}

		fmt.Printf("P[%d]: %q%s%s\n", i, displayText, flags, fontInfo)
	}
}
