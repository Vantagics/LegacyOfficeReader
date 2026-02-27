package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find the last row (row 3)
			tblStart := strings.Index(content, "<w:tbl>")
			tblEnd := strings.Index(content, "</w:tbl>")
			if tblStart < 0 || tblEnd < 0 {
				return
			}
			tblContent := content[tblStart:tblEnd+8]

			// Find the last <w:tr>
			lastTr := strings.LastIndex(tblContent, "<w:tr>")
			if lastTr < 0 {
				return
			}
			lastTrEnd := strings.Index(tblContent[lastTr:], "</w:tr>")
			if lastTrEnd < 0 {
				return
			}
			lastRow := tblContent[lastTr:lastTr+lastTrEnd+7]

			// Find the 3rd cell (index 2) - the description cell
			cellIdx := 0
			searchIdx := 0
			for cellIdx < 3 {
				tcStart := strings.Index(lastRow[searchIdx:], "<w:tc>")
				if tcStart < 0 {
					break
				}
				tcStart += searchIdx
				tcEnd := strings.Index(lastRow[tcStart+6:], "<w:tc>")
				if tcEnd < 0 {
					// Last cell
					tcEnd = strings.Index(lastRow[tcStart:], "</w:tc>")
					if tcEnd < 0 {
						break
					}
					tcEnd += tcStart + 7
				} else {
					tcEnd += tcStart + 6
				}

				if cellIdx == 2 {
					cell := lastRow[tcStart:tcEnd]
					fmt.Printf("Row 3, Cell 2 (description):\n%s\n", cell)
					
					// Count paragraphs in this cell
					pCount := strings.Count(cell, "<w:p>") + strings.Count(cell, "<w:p ")
					fmt.Printf("\nParagraphs in cell: %d (should be 2)\n", pCount)
				}
				cellIdx++
				searchIdx = tcStart + 6
			}
		}
	}
}
