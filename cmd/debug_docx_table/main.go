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
		fmt.Println("Error:", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			
			content := string(data)
			
			// Find the table
			tblStart := strings.Index(content, "<w:tbl>")
			tblEnd := strings.Index(content, "</w:tbl>")
			if tblStart >= 0 && tblEnd >= 0 {
				tbl := content[tblStart : tblEnd+len("</w:tbl>")]
				
				// Count rows
				rowCount := strings.Count(tbl, "<w:tr>")
				fmt.Printf("Table has %d rows\n\n", rowCount)
				
				// Extract each row
				remaining := tbl
				rowNum := 0
				for {
					trStart := strings.Index(remaining, "<w:tr>")
					if trStart < 0 {
						break
					}
					trEnd := strings.Index(remaining[trStart:], "</w:tr>")
					if trEnd < 0 {
						break
					}
					row := remaining[trStart : trStart+trEnd+len("</w:tr>")]
					remaining = remaining[trStart+trEnd+len("</w:tr>"):]
					rowNum++
					
					// Count cells
					cellCount := strings.Count(row, "<w:tc>")
					fmt.Printf("Row %d: %d cells\n", rowNum, cellCount)
					
					// Extract cell text
					cellRemaining := row
					cellNum := 0
					for {
						tcStart := strings.Index(cellRemaining, "<w:tc>")
						if tcStart < 0 {
							break
						}
						tcEnd := strings.Index(cellRemaining[tcStart:], "</w:tc>")
						if tcEnd < 0 {
							break
						}
						cell := cellRemaining[tcStart : tcStart+tcEnd+len("</w:tc>")]
						cellRemaining = cellRemaining[tcStart+tcEnd+len("</w:tc>"):]
						cellNum++
						
						// Extract text from <w:t> elements
						var texts []string
						cellContent := cell
						for {
							tStart := strings.Index(cellContent, "<w:t")
							if tStart < 0 {
								break
							}
							// Find the closing >
							gtPos := strings.Index(cellContent[tStart:], ">")
							if gtPos < 0 {
								break
							}
							tEnd := strings.Index(cellContent[tStart+gtPos:], "</w:t>")
							if tEnd < 0 {
								break
							}
							text := cellContent[tStart+gtPos+1 : tStart+gtPos+tEnd]
							texts = append(texts, text)
							cellContent = cellContent[tStart+gtPos+tEnd+len("</w:t>"):]
						}
						
						// Count paragraphs in cell
						pCount := strings.Count(cell, "<w:p>") + strings.Count(cell, "<w:p/>")
						
						textStr := strings.Join(texts, " | ")
						if len(textStr) > 60 {
							textStr = textStr[:60] + "..."
						}
						fmt.Printf("  Cell %d (%d paras): %q\n", cellNum, pCount, textStr)
					}
				}
			}
		}
	}
}
