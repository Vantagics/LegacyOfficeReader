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

	// Read DOCX output
	ourDoc := readZipEntry("testfie/test_new9.docx", "word/document.xml")
	ourParas := splitParas(ourDoc)

	// Build DOC text list (non-table, non-empty meaningful paragraphs)
	fmt.Println("=== Comparing DOC paragraphs to DOCX output ===")

	// Extract all text from DOCX paragraphs
	var docxTexts []string
	for _, p := range ourParas {
		if strings.Contains(p, "<w:tbl") {
			// Extract table cell texts
			cells := extractTableCellTexts(p)
			for _, c := range cells {
				docxTexts = append(docxTexts, c)
			}
		} else {
			docxTexts = append(docxTexts, extractText(p))
		}
	}

	// Extract all text from DOC paragraphs
	var docTexts []string
	for _, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		// Clean: remove special chars
		cleaned := cleanDocText(text)
		docTexts = append(docTexts, cleaned)
	}

	// Find mismatches
	fmt.Printf("DOC text entries: %d\n", len(docTexts))
	fmt.Printf("DOCX text entries: %d\n", len(docxTexts))

	// Simple sequential comparison - find first mismatch
	di, xi := 0, 0
	mismatches := 0
	for di < len(docTexts) && xi < len(docxTexts) {
		dt := docTexts[di]
		xt := docxTexts[xi]

		// Normalize for comparison
		dtNorm := normalizeText(dt)
		xtNorm := normalizeText(xt)

		if dtNorm == xtNorm {
			di++
			xi++
			continue
		}

		// Check if DOC text is empty (might be skipped in DOCX)
		if dtNorm == "" {
			di++
			continue
		}

		// Check if DOCX text is empty
		if xtNorm == "" {
			xi++
			continue
		}

		mismatches++
		if mismatches <= 15 {
			fmt.Printf("\nMISMATCH at DOC[%d] vs DOCX[%d]:\n", di, xi)
			if len(dt) > 100 {
				fmt.Printf("  DOC:  %q...\n", dt[:100])
			} else {
				fmt.Printf("  DOC:  %q\n", dt)
			}
			if len(xt) > 100 {
				fmt.Printf("  DOCX: %q...\n", xt[:100])
			} else {
				fmt.Printf("  DOCX: %q\n", xt)
			}
		}
		di++
		xi++
	}

	fmt.Printf("\nTotal mismatches: %d\n", mismatches)
}

func cleanDocText(s string) string {
	var runes []rune
	for _, r := range s {
		if r == 0x01 || r == 0x02 || r == 0x03 || r == 0x04 || r == 0x05 || r == 0x07 || r == 0x08 || r == 0x0C || r == 0x0D {
			continue
		}
		runes = append(runes, r)
	}
	return string(runes)
}

func normalizeText(s string) string {
	s = strings.ReplaceAll(s, "\t", "")
	s = strings.ReplaceAll(s, " ", "")
	s = strings.ReplaceAll(s, "\u00a0", "")
	s = strings.ReplaceAll(s, "&#x9;", "")
	s = strings.ReplaceAll(s, "&amp;", "&")
	s = strings.TrimSpace(s)
	return s
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

func extractTableCellTexts(xml string) []string {
	var texts []string
	rest := xml
	for {
		idx := strings.Index(rest, "<w:tc")
		if idx < 0 {
			break
		}
		endIdx := strings.Index(rest[idx:], "</w:tc>")
		if endIdx < 0 {
			break
		}
		cell := rest[idx : idx+endIdx+7]
		texts = append(texts, extractText(cell))
		rest = rest[idx+endIdx+7:]
	}
	return texts
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
