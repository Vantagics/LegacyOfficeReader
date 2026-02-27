package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	path := "testfie/test_v2.docx"
	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Count paragraphs
			pCount := strings.Count(content, "<w:p>") + strings.Count(content, "<w:p/>")
			fmt.Printf("Total paragraphs: %d\n", pCount)

			// Count tables
			tblCount := strings.Count(content, "<w:tbl>")
			fmt.Printf("Total tables: %d\n", tblCount)

			// Count images
			imgCount := strings.Count(content, "<w:drawing>")
			fmt.Printf("Total images/drawings: %d\n", imgCount)

			// Count page breaks
			pbCount := strings.Count(content, `w:type="page"`)
			fmt.Printf("Page breaks: %d\n", pbCount)

			// Count text boxes
			txbxCount := strings.Count(content, "<wps:txbx>")
			fmt.Printf("Text boxes: %d\n", txbxCount)

			// Count TOC fields
			tocCount := strings.Count(content, "TOC")
			fmt.Printf("TOC references: %d\n", tocCount)

			// Check for section breaks inside paragraphs
			inlineSectPr := strings.Count(content, "<w:pPr><w:sectPr") +
				strings.Count(content, "</w:jc><w:sectPr") +
				strings.Count(content, "</w:numPr><w:sectPr")
			fmt.Printf("Inline sectPr (in pPr): ~%d\n", inlineSectPr)

			// Check headings
			for level := 1; level <= 4; level++ {
				hCount := strings.Count(content, fmt.Sprintf(`w:val="Heading%d"`, level))
				if hCount > 0 {
					fmt.Printf("Heading%d count: %d\n", level, hCount)
				}
			}

			// Print first 20 paragraphs text content
			fmt.Println("\n=== First 30 paragraphs ===")
			idx := 0
			pNum := 0
			for pNum < 30 {
				pos := strings.Index(content[idx:], "<w:p>")
				if pos < 0 {
					break
				}
				pStart := idx + pos
				pEnd := strings.Index(content[pStart:], "</w:p>")
				if pEnd < 0 {
					break
				}
				pContent := content[pStart : pStart+pEnd+len("</w:p>")]

				// Extract text
				var texts []string
				tIdx := 0
				for {
					tPos := strings.Index(pContent[tIdx:], "<w:t")
					if tPos < 0 {
						break
					}
					tStart := tIdx + tPos
					// Find > after <w:t
					gtPos := strings.Index(pContent[tStart:], ">")
					if gtPos < 0 {
						break
					}
					textStart := tStart + gtPos + 1
					textEnd := strings.Index(pContent[textStart:], "</w:t>")
					if textEnd < 0 {
						break
					}
					text := pContent[textStart : textStart+textEnd]
					if text != "" {
						texts = append(texts, text)
					}
					tIdx = textStart + textEnd + len("</w:t>")
				}

				hasDrawing := strings.Contains(pContent, "<w:drawing>")
				hasTxbx := strings.Contains(pContent, "<wps:txbx>")
				hasBreak := strings.Contains(pContent, `w:type="page"`)
				hasSectPr := strings.Contains(pContent, "<w:sectPr>")

				flags := ""
				if hasDrawing {
					flags += " [IMG]"
				}
				if hasTxbx {
					flags += " [TXBX]"
				}
				if hasBreak {
					flags += " [PAGEBREAK]"
				}
				if hasSectPr {
					flags += " [SECTPR]"
				}

				text := strings.Join(texts, "")
				if len(text) > 80 {
					text = text[:80] + "..."
				}
				fmt.Printf("P[%d]: %q%s\n", pNum, text, flags)

				idx = pStart + pEnd + len("</w:p>")
				pNum++
			}
		}
	}
}
