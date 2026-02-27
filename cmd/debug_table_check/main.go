package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error:", err)
		return
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
				table := content[tblStart : tblEnd+len("</w:tbl>")]
				// Count rows
				rows := strings.Count(table, "<w:tr>")
				cells := strings.Count(table, "<w:tc>")
				fmt.Printf("Table: %d rows, %d cells\n", rows, cells)
				
				// Show first row
				firstRow := strings.Index(table, "<w:tr>")
				firstRowEnd := strings.Index(table[firstRow+5:], "</w:tr>")
				if firstRow >= 0 && firstRowEnd >= 0 {
					row := table[firstRow : firstRow+5+firstRowEnd+len("</w:tr>")]
					if len(row) > 1000 { row = row[:1000] + "..." }
					fmt.Printf("\nFirst row:\n%s\n", row)
				}
			}
		}
	}
}
