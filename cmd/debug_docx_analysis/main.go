package main

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	// Analyze the source DOC
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening DOC: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	images := d.GetImages()
	fmt.Printf("=== DOC Analysis ===\n")
	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))
	fmt.Printf("Total images: %d\n", len(images))
	fmt.Printf("Headers: %v\n", fc.Headers)
	fmt.Printf("Footers: %v\n", fc.Footers)

	// Show first 30 paragraphs in detail (title page area)
	fmt.Printf("\n=== First 30 Paragraphs (Title Page) ===\n")
	for i := 0; i < 30 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.HasPageBreak {
			flags += " PAGEBREAK"
		}
		if p.PageBreakBefore {
			flags += " PBB"
		}
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" SECT(%d)", p.SectionType)
		}
		if p.InTable {
			flags += " TABLE"
		}
		if p.TableRowEnd {
			flags += " ROWEND"
		}
		if p.IsListItem {
			flags += fmt.Sprintf(" LIST(%d,L%d)", p.ListType, p.ListLevel)
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN%v", p.DrawnImages)
		}
		// Check for inline images
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				flags += fmt.Sprintf(" INLINE_IMG[%d]", r.ImageRef)
			}
			if strings.Contains(r.Text, "\x01") {
				flags += " HAS_0x01"
			}
		}
		align := ""
		switch p.Props.Alignment {
		case 1:
			align = "center"
		case 2:
			align = "right"
		case 3:
			align = "both"
		}
		if align != "" {
			flags += " align=" + align
		}

		// Show run details
		runInfo := ""
		for ri, r := range p.Runs {
			if ri > 0 {
				runInfo += " | "
			}
			rFlags := ""
			if r.Props.Bold {
				rFlags += "B"
			}
			if r.Props.Italic {
				rFlags += "I"
			}
			if r.Props.FontSize > 0 {
				rFlags += fmt.Sprintf(" sz=%d", r.Props.FontSize)
			}
			if r.Props.FontName != "" {
				rFlags += " " + r.Props.FontName
			}
			preview := r.Text
			if len(preview) > 30 {
				preview = preview[:30] + "..."
			}
			runInfo += fmt.Sprintf("[%q%s]", preview, rFlags)
		}

		preview := text
		if len(preview) > 60 {
			preview = preview[:60] + "..."
		}
		fmt.Printf("P%d: %q%s\n", i, preview, flags)
		if len(p.Runs) > 1 || (len(p.Runs) == 1 && (p.Runs[0].Props.Bold || p.Runs[0].Props.FontSize > 30)) {
			fmt.Printf("    Runs: %s\n", runInfo)
		}
	}

	// Show page breaks and section breaks
	fmt.Printf("\n=== Page/Section Breaks ===\n")
	for i, p := range fc.Paragraphs {
		if p.HasPageBreak || p.IsSectionBreak || p.PageBreakBefore {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 40 {
				text = text[:40] + "..."
			}
			fmt.Printf("P%d: %q pageBreak=%v sectionBreak=%v(%d) pbb=%v\n",
				i, text, p.HasPageBreak, p.IsSectionBreak, p.SectionType, p.PageBreakBefore)
		}
	}

	// Show all paragraphs with images
	fmt.Printf("\n=== Paragraphs with Images ===\n")
	for i, p := range fc.Paragraphs {
		hasImg := len(p.DrawnImages) > 0
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				hasImg = true
			}
		}
		if hasImg {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			fmt.Printf("P%d: %q drawn=%v", i, text, p.DrawnImages)
			for _, r := range p.Runs {
				if r.ImageRef >= 0 {
					fmt.Printf(" inline=%d", r.ImageRef)
				}
			}
			fmt.Println()
		}
	}

	// Show image details
	fmt.Printf("\n=== Image Details ===\n")
	for i, img := range images {
		fmt.Printf("BSE[%d]: format=%d size=%d bytes\n", i, img.Format, len(img.Data))
	}

	// Analyze the output DOCX
	fmt.Printf("\n=== DOCX Analysis ===\n")
	zr, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening DOCX: %v\n", err)
		return
	}
	defer zr.Close()

	for _, f := range zr.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}
}
