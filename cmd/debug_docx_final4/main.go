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

	document, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	fc := document.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Show paragraphs around the title page (P[0]-P[50])
	fmt.Println("=== Title Page Area (P[0]-P[50]) ===")
	for i := 0; i < len(fc.Paragraphs) && i <= 50; i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		// Clean for display
		text = strings.ReplaceAll(text, "\x01", "[IMG]")
		text = strings.ReplaceAll(text, "\x08", "[DRW]")
		text = strings.ReplaceAll(text, "\x0C", "[PB]")
		if len(text) > 100 {
			text = text[:100] + "..."
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
		flags := ""
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" SECBRK(%d)", p.SectionType)
		}
		if p.HasPageBreak {
			flags += " PGBRK"
		}
		if p.TextBoxText != "" {
			flags += " TXBOX"
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN%v", p.DrawnImages)
		}
		if p.InTable {
			flags += " TBL"
		}
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}
		fmt.Printf("P[%d] %s%s: %q\n", i, align, flags, text)
	}

	// Show paragraphs around the revision table (P[50]-P[135])
	fmt.Println("\n=== Revision Table Area ===")
	for i := 50; i < len(fc.Paragraphs) && i <= 135; i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.ReplaceAll(text, "\x01", "[IMG]")
		text = strings.ReplaceAll(text, "\x08", "[DRW]")
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		flags := ""
		if p.InTable {
			flags += " TBL"
		}
		if p.TableRowEnd {
			flags += " ROWEND"
		}
		if p.IsTableCellEnd {
			flags += " CELLEND"
		}
		if p.HasPageBreak {
			flags += " PGBRK"
		}
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}
		if text != "" || flags != "" {
			fmt.Printf("P[%d]%s: %q\n", i, flags, text)
		}
	}

	// Show paragraphs around deployment diagrams (P[210]-P[279])
	fmt.Println("\n=== Deployment Diagram Area ===")
	for i := 210; i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.ReplaceAll(text, "\x01", "[IMG]")
		text = strings.ReplaceAll(text, "\x08", "[DRW]")
		if len(text) > 100 {
			text = text[:100] + "..."
		}
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN%v", p.DrawnImages)
		}
		if p.HasPageBreak {
			flags += " PGBRK"
		}
		if p.IsListItem {
			flags += " LIST"
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
		fmt.Printf("P[%d] %s%s: %q\n", i, align, flags, text)
	}

	// Show run details for specific paragraphs with formatting
	fmt.Println("\n=== Run details for key paragraphs ===")
	keyParas := []int{135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145}
	for _, idx := range keyParas {
		if idx >= len(fc.Paragraphs) {
			continue
		}
		p := fc.Paragraphs[idx]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		fmt.Printf("\nP[%d]: %q\n", idx, text)
		for j, r := range p.Runs {
			rText := r.Text
			if len(rText) > 60 {
				rText = rText[:60] + "..."
			}
			fmt.Printf("  Run[%d]: font=%q sz=%d bold=%v color=%q: %q\n",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Color, rText)
		}
	}
}
