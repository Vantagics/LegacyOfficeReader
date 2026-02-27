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

			// Find TOC area - between fldChar begin and end
			tocStart := strings.Index(content, "TOC \\o")
			if tocStart < 0 {
				fmt.Println("No TOC found")
				return
			}
			
			// Find the separate marker
			sepIdx := strings.Index(content[tocStart:], `fldCharType="separate"`)
			if sepIdx < 0 {
				fmt.Println("No TOC separate marker")
				return
			}
			sepIdx += tocStart
			
			// Find the end marker
			endIdx := strings.Index(content[sepIdx:], `fldCharType="end"`)
			if endIdx < 0 {
				fmt.Println("No TOC end marker")
				return
			}
			endIdx += sepIdx
			
			tocContent := content[sepIdx:endIdx]
			
			// Count TOC paragraphs
			tocPCount := strings.Count(tocContent, "<w:p>") + strings.Count(tocContent, "<w:p ")
			fmt.Printf("TOC paragraphs: %d\n", tocPCount)
			
			// Check for tab characters in TOC
			tabCount := strings.Count(tocContent, "<w:tab/>")
			fmt.Printf("Tab characters in TOC: %d\n", tabCount)
			
			// Check for TOC styles
			for i := 1; i <= 3; i++ {
				count := strings.Count(tocContent, fmt.Sprintf(`w:val="TOC%d"`, i))
				fmt.Printf("TOC%d style refs: %d\n", i, count)
			}
			
			// Check for key TOC text
			tocTexts := []string{"引言", "产品概述", "产品组成与架构", "威胁情报", "分析平台", "典型部署"}
			for _, t := range tocTexts {
				if strings.Contains(tocContent, t) {
					fmt.Printf("  OK: %s\n", t)
				} else {
					fmt.Printf("  MISSING: %s\n", t)
				}
			}
			
			// Show first TOC paragraph
			fmt.Println("\n=== First TOC entry ===")
			pStart := strings.Index(tocContent, "<w:p>")
			if pStart < 0 {
				pStart = strings.Index(tocContent, "<w:p ")
			}
			if pStart >= 0 {
				pEnd := strings.Index(tocContent[pStart:], "</w:p>")
				if pEnd >= 0 {
					fmt.Println(tocContent[pStart:pStart+pEnd+6])
				}
			}
		}
	}
}
