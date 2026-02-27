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

	// Check specific paragraphs for formatting details
	// Title page paragraphs (P[0]-P[20])
	fmt.Println("=== Title Page Paragraphs ===")
	for i := 0; i < 25 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.ReplaceAll(text, "\t", "\\t")
		text = strings.ReplaceAll(text, "\r", "\\r")
		if len(text) > 60 {
			text = text[:60] + "..."
		}
		align := []string{"left", "center", "right", "both"}[p.Props.Alignment]
		fmt.Printf("P[%d] align=%s indent=(%d,%d,%d) space=(%d,%d) line=(%d,%d)",
			i, align, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule)
		if p.IsSectionBreak {
			fmt.Printf(" SECBRK(%d)", p.SectionType)
		}
		if p.HasPageBreak {
			fmt.Printf(" PAGEBRK")
		}
		if len(p.DrawnImages) > 0 {
			fmt.Printf(" DRAWN%v", p.DrawnImages)
		}
		if p.TextBoxText != "" {
			fmt.Printf(" TXBX")
		}
		// Show run formatting
		for j, r := range p.Runs {
			if j > 2 {
				break
			}
			fmt.Printf("\n  R[%d] font=%q sz=%d bold=%v italic=%v color=%s",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Italic, r.Props.Color)
		}
		fmt.Printf("\n  text=%q\n", text)
	}

	// Check a few body paragraphs
	fmt.Println("\n=== Sample Body Paragraphs ===")
	for _, idx := range []int{135, 136, 137, 152, 153, 154, 155, 159, 172, 195, 213, 214, 215} {
		if idx >= len(fc.Paragraphs) {
			continue
		}
		p := fc.Paragraphs[idx]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.ReplaceAll(text, "\t", "\\t")
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		align := []string{"left", "center", "right", "both"}[p.Props.Alignment]
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.IsListItem {
			flags += fmt.Sprintf(" LIST(t=%d,l=%d)", p.ListType, p.ListLevel)
		}
		fmt.Printf("P[%d] align=%s%s indent=(%d,%d,%d) space=(%d,%d) line=(%d,%d)\n",
			idx, align, flags, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule)
		for j, r := range p.Runs {
			if j > 1 {
				break
			}
			rtext := r.Text
			if len(rtext) > 40 {
				rtext = rtext[:40] + "..."
			}
			fmt.Printf("  R[%d] font=%q sz=%d bold=%v: %q\n",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, rtext)
		}
	}

	// Check table paragraphs
	fmt.Println("\n=== Table Paragraphs ===")
	tableStart := -1
	for i, p := range fc.Paragraphs {
		if p.InTable && tableStart < 0 {
			tableStart = i
		}
		if p.InTable {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			text = strings.ReplaceAll(text, "\t", "\\t")
			text = strings.ReplaceAll(text, "\x07", "")
			if len(text) > 50 {
				text = text[:50] + "..."
			}
			flags := ""
			if p.TableRowEnd {
				flags += " ROWEND"
				if len(p.CellWidths) > 0 {
					flags += fmt.Sprintf(" widths=%v", p.CellWidths)
				}
			}
			if p.IsTableCellEnd {
				flags += " CELLEND"
			}
			fmt.Printf("P[%d]%s: %q\n", i, flags, text)
		}
	}

	// Check drawn image paragraphs
	fmt.Println("\n=== Drawn Image Paragraphs ===")
	for i, p := range fc.Paragraphs {
		if len(p.DrawnImages) > 0 {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 40 {
				text = text[:40] + "..."
			}
			fmt.Printf("P[%d] drawn=%v text=%q\n", i, p.DrawnImages, text)
		}
	}

	// Check inline image paragraphs
	fmt.Println("\n=== Inline Image Paragraphs ===")
	for i, p := range fc.Paragraphs {
		for j, r := range p.Runs {
			if r.ImageRef >= 0 {
				text := ""
				for _, rr := range p.Runs {
					text += rr.Text
				}
				if len(text) > 40 {
					text = text[:40] + "..."
				}
				fmt.Printf("P[%d] R[%d] imageRef=%d text=%q\n", i, j, r.ImageRef, text)
				break
			}
		}
	}
}
