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

			// Find the paragraph containing "目"
			idx := strings.Index(content, ">目 ")
			if idx >= 0 {
				// Find the start of this paragraph
				pStart := strings.LastIndex(content[:idx], "<w:p>")
				if pStart < 0 {
					pStart = strings.LastIndex(content[:idx], "<w:p ")
				}
				// Find the end of this paragraph
				pEnd := strings.Index(content[idx:], "</w:p>")
				if pEnd >= 0 {
					pEnd += idx + 6
				}
				if pStart >= 0 && pEnd > pStart {
					fmt.Printf("TOC heading paragraph:\n%s\n", content[pStart:pEnd])
				}
			}

			// Also check: is "目 录" present as a combined string?
			// The original doc has "目 " + " " + "录" as 3 runs
			// Check if the text "目" and "录" are in the same paragraph
			fmt.Println("\n--- Searching for 录 near 目 ---")
			idx2 := strings.Index(content, ">目 ")
			if idx2 >= 0 {
				end := idx2 + 300
				if end > len(content) {
					end = len(content)
				}
				snippet := content[idx2:end]
				if strings.Contains(snippet, "录") {
					fmt.Println("OK: 录 found near 目 in same area")
				} else {
					fmt.Println("ISSUE: 录 not found near 目")
				}
			}
		}
	}
}
