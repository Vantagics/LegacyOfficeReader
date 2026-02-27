package main

import (
	"fmt"
	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}
	
	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))
	fmt.Printf("Header entries: %d\n", len(fc.HeaderEntries))
	fmt.Printf("Footer entries: %d\n", len(fc.FooterEntries))
	
	for i, he := range fc.HeaderEntries {
		fmt.Printf("  Header[%d]: type=%s images=%v text=%q\n", i, he.Type, he.Images, truncate(he.Text, 60))
	}
	for i, fe := range fc.FooterEntries {
		fmt.Printf("  Footer[%d]: type=%s images=%v text=%q rawText=%q\n", i, fe.Type, fe.Images, truncate(fe.Text, 60), truncate(fe.RawText, 80))
	}
	
	fmt.Println("\n--- Page structure analysis ---")
	pageBreakCount := 0
	sectionBreakCount := 0
	imageCount := 0
	textboxCount := 0
	tableCount := 0
	
	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		
		// Show interesting paragraphs
		interesting := false
		reason := ""
		if p.IsSectionBreak {
			sectionBreakCount++
			interesting = true
			reason += fmt.Sprintf("SECTION_BREAK(type=%d) ", p.SectionType)
		}
		if p.HasPageBreak || p.PageBreakBefore {
			pageBreakCount++
			interesting = true
			reason += "PAGE_BREAK "
		}
		if len(p.DrawnImages) > 0 {
			imageCount++
			interesting = true
			reason += fmt.Sprintf("DRAWN=%v ", p.DrawnImages)
		}
		if p.TextBoxText != "" {
			textboxCount++
			interesting = true
			reason += fmt.Sprintf("TEXTBOX=%q ", truncate(p.TextBoxText, 40))
		}
		if p.InTable {
			if p.TableRowEnd {
				tableCount++
			}
			interesting = true
			reason += "TABLE "
		}
		if p.HeadingLevel > 0 {
			interesting = true
			reason += fmt.Sprintf("H%d ", p.HeadingLevel)
		}
		
		if interesting {
			fmt.Printf("P[%d] %s align=%d text=%q\n", i, reason, p.Props.Alignment, truncate(text, 60))
		}
	}
	
	fmt.Printf("\nSummary: %d sections, %d page breaks, %d drawn images, %d textboxes, %d table rows\n",
		sectionBreakCount, pageBreakCount, imageCount, textboxCount, tableCount)
}

func truncate(s string, n int) string {
	runes := []rune(s)
	if len(runes) > n {
		return string(runes[:n]) + "..."
	}
	return s
}
