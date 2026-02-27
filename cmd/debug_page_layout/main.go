package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	files := []string{"testfie/test_new8.docx"}

	for _, f := range files {
		fmt.Printf("\n=== %s ===\n", f)
		doc := readZipEntry(f, "word/document.xml")
		paras := splitParas(doc)

		fmt.Printf("Total XML elements: %d\n\n", len(paras))

		for i, p := range paras {
			text := extractText(p)
			runes := []rune(text)

			// Check for page breaks
			hasPageBreak := strings.Contains(p, "w:type=\"page\"")
			hasSectBreak := strings.Contains(p, "<w:sectPr")
			hasLastRendered := strings.Contains(p, "lastRenderedPageBreak")

			// Check for images
			hasInlineImg := strings.Contains(p, "<wp:inline")
			hasAnchorImg := strings.Contains(p, "<wp:anchor")
			isTable := strings.Contains(p, "<w:tbl")

			// Get style
			style := ""
			if idx := strings.Index(p, "pStyle"); idx >= 0 {
				sub := p[idx:]
				valIdx := strings.Index(sub, "w:val=\"")
				if valIdx >= 0 {
					end := strings.Index(sub[valIdx+7:], "\"")
					if end >= 0 {
						style = sub[valIdx+7 : valIdx+7+end]
					}
				}
			}

			// Only show interesting paragraphs
			interesting := hasPageBreak || hasSectBreak || hasLastRendered ||
				hasInlineImg || hasAnchorImg || isTable ||
				style != "" || len(runes) == 0

			if !interesting && len(runes) < 40 {
				interesting = true
			}

			flags := ""
			if hasPageBreak {
				flags += " [PAGE_BREAK]"
			}
			if hasSectBreak {
				flags += " [SECT_BREAK]"
			}
			if hasLastRendered {
				flags += " [LAST_RENDERED_PB]"
			}
			if hasInlineImg {
				flags += " [INLINE_IMG]"
			}
			if hasAnchorImg {
				flags += " [ANCHOR_IMG]"
			}
			if isTable {
				flags += " [TABLE]"
			}
			if style != "" {
				flags += " [" + style + "]"
			}
			if len(runes) == 0 && !isTable {
				flags += " [EMPTY]"
			}

			preview := string(runes)
			if len(runes) > 80 {
				preview = string(runes[:80]) + "..."
			}

			fmt.Printf("  [%3d] chars=%3d%s %s\n", i, len(runes), flags, preview)
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
