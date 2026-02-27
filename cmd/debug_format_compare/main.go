package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "parse: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("no formatted content")
		return
	}

	// Read DOCX output for comparison
	ourDoc := readZipEntry("testfie/test_new9.docx", "word/document.xml")

	// Show DOC paragraph formatting for key paragraphs
	fmt.Println("=== DOC Paragraph Formatting (key paragraphs) ===")
	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		runes := []rune(text)

		// Show formatting for non-empty, non-table paragraphs
		if len(runes) == 0 && len(p.DrawnImages) == 0 {
			continue
		}
		if p.InTable {
			continue
		}

		preview := string(runes)
		if len(runes) > 40 {
			preview = string(runes[:40]) + "..."
		}

		fmt.Printf("\nDOC[%3d] %q\n", i, preview)
		fmt.Printf("  Align=%d(set=%v) IndL=%d IndR=%d IndFirst=%d\n",
			p.Props.Alignment, p.Props.AlignmentSet,
			p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst)
		fmt.Printf("  SpBefore=%d SpAfter=%d LineSpacing=%d LineRule=%d\n",
			p.Props.SpaceBefore, p.Props.SpaceAfter,
			p.Props.LineSpacing, p.Props.LineRule)
		fmt.Printf("  Heading=%d List=%v ListType=%d PageBreak=%v SectBreak=%v\n",
			p.HeadingLevel, p.IsListItem, p.ListType, p.HasPageBreak, p.IsSectionBreak)

		// Show run formatting
		for j, r := range p.Runs {
			if j > 2 {
				fmt.Printf("  ... (%d more runs)\n", len(p.Runs)-3)
				break
			}
			fmt.Printf("  Run[%d]: font=%q sz=%d bold=%v italic=%v color=%q\n",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Italic, r.Props.Color)
		}
	}

	// Check DOCX paragraph formatting
	fmt.Println("\n\n=== DOCX Paragraph Properties (first 20 non-empty) ===")
	paras := splitParas(ourDoc)
	count := 0
	for i, p := range paras {
		if strings.Contains(p, "<w:tbl") {
			continue
		}
		text := extractText(p)
		if text == "" && !strings.Contains(p, "<wp:") {
			continue
		}
		count++
		if count > 30 {
			break
		}

		runes := []rune(text)
		preview := string(runes)
		if len(runes) > 40 {
			preview = string(runes[:40]) + "..."
		}

		// Extract pPr content
		pprStart := strings.Index(p, "<w:pPr>")
		pprEnd := strings.Index(p, "</w:pPr>")
		ppr := ""
		if pprStart >= 0 && pprEnd >= 0 {
			ppr = p[pprStart+7 : pprEnd]
		}

		// Extract first rPr
		rprStart := strings.Index(p, "<w:rPr>")
		rprEnd := strings.Index(p, "</w:rPr>")
		rpr := ""
		if rprStart >= 0 && rprEnd >= 0 {
			rpr = p[rprStart+7 : rprEnd]
		}

		fmt.Printf("\nXML[%3d] %q\n", i, preview)
		fmt.Printf("  pPr: %s\n", ppr)
		if rpr != "" {
			fmt.Printf("  rPr: %s\n", rpr)
		}
	}
}

func extractText(xml string) string {
	var sb strings.Builder
	rest := xml
	for {
		idx := strings.Index(rest, "<w:t")
		if idx < 0 {
			break
		}
		gt := strings.Index(rest[idx:], ">")
		if gt < 0 {
			break
		}
		start := idx + gt + 1
		end := strings.Index(rest[start:], "</w:t>")
		if end < 0 {
			break
		}
		sb.WriteString(rest[start : start+end])
		rest = rest[start+end+6:]
	}
	return sb.String()
}

func readZipEntry(zipPath, entry string) string {
	r, err := zip.OpenReader(zipPath)
	if err != nil {
		return ""
	}
	defer r.Close()
	for _, f := range r.File {
		if f.Name == entry {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			return string(data)
		}
	}
	return ""
}

func splitParas(xml string) []string {
	var result []string
	rest := xml
	for {
		pIdx := strings.Index(rest, "<w:p ")
		p2Idx := strings.Index(rest, "<w:p>")
		tIdx := strings.Index(rest, "<w:tbl")
		minIdx := -1
		if pIdx >= 0 {
			minIdx = pIdx
		}
		if p2Idx >= 0 && (minIdx < 0 || p2Idx < minIdx) {
			minIdx = p2Idx
		}
		if tIdx >= 0 && (minIdx < 0 || tIdx < minIdx) {
			minIdx = tIdx
		}
		if minIdx < 0 {
			break
		}
		if minIdx == tIdx {
			endTag := "</w:tbl>"
			endIdx := strings.Index(rest[minIdx:], endTag)
			if endIdx < 0 {
				result = append(result, rest[minIdx:])
				break
			}
			result = append(result, rest[minIdx:minIdx+endIdx+len(endTag)])
			rest = rest[minIdx+endIdx+len(endTag):]
		} else {
			endTag := "</w:p>"
			endIdx := strings.Index(rest[minIdx:], endTag)
			if endIdx < 0 {
				result = append(result, rest[minIdx:])
				break
			}
			result = append(result, rest[minIdx:minIdx+endIdx+len(endTag)])
			rest = rest[minIdx+endIdx+len(endTag):]
		}
	}
	return result
}
