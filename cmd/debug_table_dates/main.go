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

			// Check for date components
			dates := []string{"2021", "2022", "7.20", "8.30", "9.29"}
			for _, d := range dates {
				if strings.Contains(content, d) {
					fmt.Printf("  OK: %s found\n", d)
				} else {
					fmt.Printf("  MISSING: %s\n", d)
				}
			}

			// Check for "版本号" - might be split as "版本" + "号"
			if strings.Contains(content, "版本号") {
				fmt.Println("OK: 版本号 (combined)")
			} else if strings.Contains(content, "版本") {
				fmt.Println("OK: 版本 found (号 might be in next run)")
			}

			// Check for "修订记录"
			if strings.Contains(content, "修订记录") {
				fmt.Println("OK: 修订记录")
			} else if strings.Contains(content, "修订") {
				fmt.Println("PARTIAL: 修订 found")
			} else {
				fmt.Println("MISSING: 修订记录")
			}

			// Check for "3.0.10.0" - might be split
			if strings.Contains(content, "3.0.10.0") {
				fmt.Println("OK: 3.0.10.0")
			} else if strings.Contains(content, "3.0.10") {
				fmt.Println("PARTIAL: 3.0.10 found")
			} else if strings.Contains(content, "0.10") {
				fmt.Println("PARTIAL: 0.10 found")
			} else {
				fmt.Println("MISSING: 3.0.10.0 completely")
			}

			// Find the table header row content
			tblStart := strings.Index(content, "<w:tbl>")
			if tblStart >= 0 {
				// Find first row
				trStart := strings.Index(content[tblStart:], "<w:tr>")
				if trStart >= 0 {
					trStart += tblStart
					trEnd := strings.Index(content[trStart:], "</w:tr>")
					if trEnd >= 0 {
						row := content[trStart:trStart+trEnd+7]
						// Extract all text from this row
						var texts []string
						idx := 0
						for {
							tStart := strings.Index(row[idx:], "<w:t")
							if tStart < 0 {
								break
							}
							tStart += idx
							// Find the > after <w:t...
							gt := strings.Index(row[tStart:], ">")
							if gt < 0 {
								break
							}
							textStart := tStart + gt + 1
							tEnd := strings.Index(row[textStart:], "</w:t>")
							if tEnd < 0 {
								break
							}
							texts = append(texts, row[textStart:textStart+tEnd])
							idx = textStart + tEnd + 6
						}
						fmt.Printf("\nFirst row texts: %v\n", texts)
					}
				}
			}
		}
	}
}
