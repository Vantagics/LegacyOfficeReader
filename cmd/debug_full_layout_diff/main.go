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
	// Open DOC source
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open doc: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "parse doc: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("no formatted content")
		return
	}

	fmt.Printf("DOC paragraphs: %d\n", len(fc.Paragraphs))

	// Read our output DOCX
	ourDoc := readZipEntry("testfie/test_new8.docx", "word/document.xml")
	ourParas := splitParas(ourDoc)
	fmt.Printf("DOCX paragraphs (XML elements): %d\n", len(ourParas))

	// Show DOC paragraph details - focus on page breaks and structure
	fmt.Println("\n=== DOC Paragraph Summary ===")
	for i, p := range fc.Paragraphs {
		totalText := ""
		for _, r := range p.Runs {
			totalText += r.Text
		}
		runes := []rune(totalText)

		flags := ""
		if p.HasPageBreak {
			flags += " [PAGE_BREAK]"
		}
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" [H%d]", p.HeadingLevel)
		}
		if p.InTable {
			flags += " [TABLE]"
		}
		if p.ListType > 0 {
			flags += fmt.Sprintf(" [LIST:%d]", p.ListType)
		}
		if len(runes) == 0 {
			flags += " [EMPTY]"
		}

		// Check for images
		hasImg := false
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				hasImg = true
			}
		}
		if hasImg {
			flags += " [IMG]"
		}

		// Check for drawn images
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" [DRAWN:%d]", len(p.DrawnImages))
		}

		preview := string(runes)
		if len(runes) > 60 {
			preview = string(runes[:60]) + "..."
		}

		fmt.Printf("  DOC[%3d] runs=%d chars=%3d%s %s\n", i, len(p.Runs), len(runes), flags, preview)
	}

	// Show DOCX paragraph details
	fmt.Println("\n=== DOCX XML Paragraph Summary ===")
	for i, p := range ourParas {
		text := extractText(p)
		runes := []rune(text)

		flags := ""
		if strings.Contains(p, "<w:tbl") {
			flags += " [TABLE]"
		}
		if strings.Contains(p, "w:type=\"page\"") || strings.Contains(p, "<w:lastRenderedPageBreak/>") {
			flags += " [PAGE_BREAK]"
		}
		if strings.Contains(p, "w:val=\"Heading") || strings.Contains(p, "pStyle") {
			// Extract style
			idx := strings.Index(p, "pStyle")
			if idx >= 0 {
				sub := p[idx:]
				valIdx := strings.Index(sub, "w:val=\"")
				if valIdx >= 0 {
					end := strings.Index(sub[valIdx+7:], "\"")
					if end >= 0 {
						flags += fmt.Sprintf(" [%s]", sub[valIdx+7:valIdx+7+end])
					}
				}
			}
		}
		if strings.Contains(p, "<wp:inline") || strings.Contains(p, "<wp:anchor") {
			flags += " [IMG]"
		}
		if len(runes) == 0 && !strings.Contains(p, "<w:tbl") {
			flags += " [EMPTY]"
		}

		preview := string(runes)
		if len(runes) > 60 {
			preview = string(runes[:60]) + "..."
		}

		fmt.Printf("  XML[%3d] chars=%3d%s %s\n", i, len(runes), flags, preview)
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
