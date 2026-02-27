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

			// Find the table
			tblStart := strings.Index(content, "<w:tbl>")
			tblEnd := strings.Index(content, "</w:tbl>")
			if tblStart < 0 || tblEnd < 0 {
				fmt.Println("No table found")
				return
			}
			tblContent := content[tblStart:tblEnd+8]

			// Extract all text from each row
			rowIdx := 0
			searchIdx := 0
			for {
				trStart := strings.Index(tblContent[searchIdx:], "<w:tr>")
				if trStart < 0 {
					break
				}
				trStart += searchIdx
				trEnd := strings.Index(tblContent[trStart:], "</w:tr>")
				if trEnd < 0 {
					break
				}
				trEnd += trStart + 7
				row := tblContent[trStart:trEnd]

				// Extract text from each cell
				fmt.Printf("Row %d:\n", rowIdx)
				cellIdx := 0
				cellSearch := 0
				for {
					tcStart := strings.Index(row[cellSearch:], "<w:tc>")
					if tcStart < 0 {
						break
					}
					tcStart += cellSearch
					tcEnd := strings.Index(row[tcStart:], "</w:tc>")
					if tcEnd < 0 {
						break
					}
					tcEnd += tcStart + 7
					cell := row[tcStart:tcEnd]

					// Extract all text from this cell
					var cellTexts []string
					tIdx := 0
					for {
						tStart := strings.Index(cell[tIdx:], "<w:t")
						if tStart < 0 {
							break
						}
						tStart += tIdx
						gt := strings.Index(cell[tStart:], ">")
						if gt < 0 {
							break
						}
						textStart := tStart + gt + 1
						tEnd := strings.Index(cell[textStart:], "</w:t>")
						if tEnd < 0 {
							break
						}
						cellTexts = append(cellTexts, cell[textStart:textStart+tEnd])
						tIdx = textStart + tEnd + 6
					}
					combined := strings.Join(cellTexts, "")
					fmt.Printf("  Cell %d: %q\n", cellIdx, combined)
					cellIdx++
					cellSearch = tcEnd
				}
				rowIdx++
				searchIdx = trEnd
			}
		}
	}
}
