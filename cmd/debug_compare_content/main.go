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
		return
	}

	// Find page boundaries (page breaks and section breaks)
	fmt.Println("=== Page Boundaries ===")
	pageNum := 1
	for i, p := range fc.Paragraphs {
		if p.HasPageBreak || p.PageBreakBefore {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			text = strings.ReplaceAll(text, "\x01", "[IMG]")
			fmt.Printf("Page break at P%d (page %d->%d): %q\n", i, pageNum, pageNum+1, truncate(text, 60))
			pageNum++
		}
		if p.IsSectionBreak {
			fmt.Printf("Section break at P%d (type=%d)\n", i, p.SectionType)
		}
	}

	// Show first page content (before first page break)
	fmt.Println("\n=== Page 1 (Title Page) Content ===")
	for i, p := range fc.Paragraphs {
		if p.HasPageBreak || p.PageBreakBefore {
			fmt.Printf("--- End of page 1 at P%d ---\n", i)
			break
		}
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.ReplaceAll(text, "\x01", "[IMG]")
		text = strings.ReplaceAll(text, "\x08", "[OBJ]")
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.InTable {
			flags += " TBL"
		}
		if p.TableRowEnd {
			flags += " ROW_END"
		}
		align := []string{"left", "center", "right", "both"}[p.Props.Alignment]
		fmt.Printf("P%d%s [%s]: %q\n", i, flags, align, truncate(text, 80))
	}

	// Show page 2 content
	fmt.Println("\n=== Page 2 Content (first 20 paragraphs) ===")
	inPage2 := false
	count := 0
	for i, p := range fc.Paragraphs {
		if !inPage2 {
			if p.HasPageBreak || p.PageBreakBefore {
				inPage2 = true
			}
			continue
		}
		if p.HasPageBreak || p.PageBreakBefore {
			if count > 0 {
				break
			}
		}
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.ReplaceAll(text, "\x01", "[IMG]")
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.InTable {
			flags += " TBL"
		}
		align := []string{"left", "center", "right", "both"}[p.Props.Alignment]
		fmt.Printf("P%d%s [%s]: %q\n", i, flags, align, truncate(text, 80))
		count++
		if count >= 20 {
			break
		}
	}
}

func truncate(s string, maxLen int) string {
	runes := []rune(s)
	if len(runes) <= maxLen {
		return s
	}
	return string(runes[:maxLen]) + "..."
}
