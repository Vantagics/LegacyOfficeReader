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
			
			// Count rows and cells
			trCount := strings.Count(tblContent, "<w:tr>")
			tcCount := strings.Count(tblContent, "<w:tc>")
			fmt.Printf("Table: %d rows, %d cells\n", trCount, tcCount)
			
			// Check for key table content
			tableTexts := []string{
				"修订记录", "版本号", "修改状态", "修改内容", "修改人",
				"3.0.10.0", "3.0.11.0", "3.0.11.0.SP5",
				"2021.7.20", "2022.8.30", "2022.9.29",
				"葛成宇",
				"产品版本更新到V",
				"增加国产化内容",
			}
			for _, t := range tableTexts {
				if strings.Contains(tblContent, t) {
					fmt.Printf("  OK: %s\n", t)
				} else {
					fmt.Printf("  MISSING: %s\n", t)
				}
			}
			
			// Show first row
			fmt.Println("\n=== First row ===")
			firstTr := strings.Index(tblContent, "<w:tr>")
			if firstTr >= 0 {
				trEnd := strings.Index(tblContent[firstTr:], "</w:tr>")
				if trEnd >= 0 {
					row := tblContent[firstTr:firstTr+trEnd+7]
					if len(row) > 500 {
						row = row[:500] + "..."
					}
					fmt.Println(row)
				}
			}
		}
	}
}
