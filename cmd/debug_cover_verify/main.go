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

			// Find the first 20 paragraphs
			fmt.Println("=== First 15 paragraphs (cover page area) ===")
			searchIdx := strings.Index(content, "<w:body>") + 7
			for i := 0; i < 15; i++ {
				pStart := strings.Index(content[searchIdx:], "<w:p>")
				if pStart < 0 {
					pStart = strings.Index(content[searchIdx:], "<w:p ")
				}
				if pStart < 0 {
					break
				}
				pStart += searchIdx
				pEnd := strings.Index(content[pStart:], "</w:p>")
				if pEnd < 0 {
					break
				}
				pEnd += pStart + 6
				
				para := content[pStart:pEnd]
				
				// Summarize paragraph content
				hasText := false
				hasImage := false
				hasTextBox := false
				hasSectPr := false
				
				if strings.Contains(para, "<w:t") {
					hasText = true
				}
				if strings.Contains(para, "<pic:pic>") || strings.Contains(para, "<wp:anchor") || strings.Contains(para, "<wp:inline") {
					hasImage = true
				}
				if strings.Contains(para, "<wps:wsp>") {
					hasTextBox = true
				}
				if strings.Contains(para, "<w:sectPr") {
					hasSectPr = true
				}
				
				// Extract text
				var texts []string
				tIdx := 0
				for {
					tStart := strings.Index(para[tIdx:], "<w:t")
					if tStart < 0 {
						break
					}
					tStart += tIdx
					gt := strings.Index(para[tStart:], ">")
					if gt < 0 {
						break
					}
					textStart := tStart + gt + 1
					tEnd := strings.Index(para[textStart:], "</w:t>")
					if tEnd < 0 {
						break
					}
					texts = append(texts, para[textStart:textStart+tEnd])
					tIdx = textStart + tEnd + 6
				}
				
				textStr := strings.Join(texts, " | ")
				if len(textStr) > 80 {
					textStr = textStr[:80] + "..."
				}
				
				flags := ""
				if hasText {
					flags += " TEXT"
				}
				if hasImage {
					flags += " IMAGE"
				}
				if hasTextBox {
					flags += " TEXTBOX"
				}
				if hasSectPr {
					flags += " SECTPR"
				}
				
				fmt.Printf("P[%d]:%s %q\n", i, flags, textStr)
				searchIdx = pEnd
			}
		}
	}
	
	// Check header3 (first page header with background)
	for _, f := range r.File {
		if f.Name == "word/header3.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			
			fmt.Println("\n=== Header3 (first page) ===")
			if strings.Contains(content, "<wp:anchor") {
				fmt.Println("Has anchor image (background)")
				// Check if it's behind doc
				if strings.Contains(content, `behindDoc="1"`) {
					fmt.Println("  behindDoc=1 (background)")
				}
				// Check size
				extIdx := strings.Index(content, "<wp:extent")
				if extIdx >= 0 {
					end := extIdx + 80
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("  %s\n", content[extIdx:end])
				}
			}
			if strings.Contains(content, "<wp:inline") {
				fmt.Println("Has inline image")
			}
		}
	}
}
