package main

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	// Parse the DOC file
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL open doc: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	document, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL parse doc: %v\n", err)
		os.Exit(1)
	}

	fc := document.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	fmt.Printf("=== DOC Analysis ===\n")
	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))
	fmt.Printf("Headers: %d, Footers: %d\n", len(fc.Headers), len(fc.Footers))
	fmt.Printf("HeadersRaw: %d, FootersRaw: %d\n", len(fc.HeadersRaw), len(fc.FootersRaw))

	for i, h := range fc.Headers {
		fmt.Printf("  Header[%d]: %q\n", i, h)
	}
	for i, f := range fc.Footers {
		fmt.Printf("  Footer[%d]: %q\n", i, f)
	}
	for i, f := range fc.FootersRaw {
		fmt.Printf("  FooterRaw[%d]: %q\n", i, f)
	}

	// Show first 30 paragraphs with details
	fmt.Printf("\n=== First 30 Paragraphs ===\n")
	for i := 0; i < len(fc.Paragraphs) && i < 30; i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		// Truncate for display
		if len(text) > 80 {
			text = text[:80] + "..."
		}
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
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}
		if p.IsListItem {
			flags += " LIST"
		}
		if p.IsSectionBreak {
			flags += fmt.Sprintf(" SECBRK(%d)", p.SectionType)
		}
		if p.HasPageBreak {
			flags += " PGBRK"
		}
		if p.PageBreakBefore {
			flags += " PGBRKBEFORE"
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN%v", p.DrawnImages)
		}
		if p.TextBoxText != "" {
			flags += fmt.Sprintf(" TXBOX=%q", p.TextBoxText)
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
		if p.Props.AlignmentSet {
			align += "*"
		}
		fmt.Printf("P[%d] align=%s indent=%d/%d/%d sp=%d/%d line=%d/%d%s: %q\n",
			i, align, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule,
			flags, text)

		// Show run details for first few paragraphs
		if i < 10 {
			for j, r := range p.Runs {
				rText := r.Text
				if len(rText) > 60 {
					rText = rText[:60] + "..."
				}
				fmt.Printf("  Run[%d]: font=%q sz=%d bold=%v italic=%v color=%q img=%d: %q\n",
					j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Italic, r.Props.Color, r.ImageRef, rText)
			}
		}
	}

	// Show all headings
	fmt.Printf("\n=== All Headings ===\n")
	for i, p := range fc.Paragraphs {
		if p.HeadingLevel > 0 {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 100 {
				text = text[:100] + "..."
			}
			fmt.Printf("P[%d] H%d: %q\n", i, p.HeadingLevel, text)
		}
	}

	// Show all page breaks
	fmt.Printf("\n=== Page Breaks ===\n")
	for i, p := range fc.Paragraphs {
		if p.HasPageBreak || p.PageBreakBefore || p.IsSectionBreak {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			fmt.Printf("P[%d] pgbrk=%v pgbrkBefore=%v secbrk=%v(%d): %q\n",
				i, p.HasPageBreak, p.PageBreakBefore, p.IsSectionBreak, p.SectionType, text)
		}
	}

	// Show images info
	images := document.GetImages()
	fmt.Printf("\n=== Images: %d total ===\n", len(images))
	for i, img := range images {
		fmt.Printf("  Image[%d]: format=%d size=%d bytes\n", i, img.Format, len(img.Data))
	}

	// Show drawn image distribution
	fmt.Printf("\n=== Drawn Image Distribution ===\n")
	drawnFreq := make(map[int]int)
	for _, p := range fc.Paragraphs {
		for _, idx := range p.DrawnImages {
			drawnFreq[idx]++
		}
	}
	for idx, freq := range drawnFreq {
		fmt.Printf("  BSE[%d]: appears in %d paragraphs\n", idx, freq)
	}

	// Now inspect the generated DOCX
	fmt.Printf("\n=== DOCX Inspection ===\n")
	zr, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL open docx: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Read document.xml and show first 2000 chars
	for _, f := range zr.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			buf := make([]byte, 3000)
			n, _ := rc.Read(buf)
			rc.Close()
			content := string(buf[:n])
			fmt.Printf("\n=== document.xml (first 3000 chars) ===\n%s\n", content)
		}
		if f.Name == "word/footer1.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			buf := make([]byte, 2000)
			n, _ := rc.Read(buf)
			rc.Close()
			fmt.Printf("\n=== footer1.xml ===\n%s\n", string(buf[:n]))
		}
		if f.Name == "word/styles.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			buf := make([]byte, 5000)
			n, _ := rc.Read(buf)
			rc.Close()
			fmt.Printf("\n=== styles.xml ===\n%s\n", string(buf[:n]))
		}
	}

	// Count paragraphs in DOCX
	for _, f := range zr.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			buf := make([]byte, 500000)
			n, _ := rc.Read(buf)
			rc.Close()
			content := string(buf[:n])
			pCount := strings.Count(content, "<w:p>") + strings.Count(content, "<w:p ")
			fmt.Printf("\n=== DOCX paragraph count: %d ===\n", pCount)

			// Count images
			imgCount := strings.Count(content, "<w:drawing>")
			fmt.Printf("=== DOCX image count: %d ===\n", imgCount)

			// Count tables
			tblCount := strings.Count(content, "<w:tbl>")
			fmt.Printf("=== DOCX table count: %d ===\n", tblCount)

			// Check for section breaks
			sectCount := strings.Count(content, "<w:sectPr>")
			fmt.Printf("=== DOCX section count: %d ===\n", sectCount)
		}
	}
}
