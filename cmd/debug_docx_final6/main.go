package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find the "引言" heading and show surrounding XML
			idx := strings.Index(content, "引言")
			if idx > 0 {
				start := idx - 300
				if start < 0 {
					start = 0
				}
				end := idx + 500
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("=== Around '引言' ===\n%s\n\n", content[start:end])
			}

			// Find "产品概述" heading
			idx = strings.Index(content, "产品概述")
			if idx > 0 {
				start := idx - 300
				if start < 0 {
					start = 0
				}
				end := idx + 800
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("=== Around '产品概述' ===\n%s\n\n", content[start:end])
			}

			// Find "版权声明"
			idx = strings.Index(content, "版权声明")
			if idx > 0 {
				start := idx - 300
				if start < 0 {
					start = 0
				}
				end := idx + 500
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("=== Around '版权声明' ===\n%s\n\n", content[start:end])
			}

			// Find "创建时间"
			idx = strings.Index(content, "创建时间")
			if idx > 0 {
				start := idx - 300
				if start < 0 {
					start = 0
				}
				end := idx + 500
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("=== Around '创建时间' ===\n%s\n\n", content[start:end])
			}

			// Find first table
			idx = strings.Index(content, "<w:tbl>")
			if idx > 0 {
				end := idx + 2000
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("=== Table start ===\n%s\n\n", content[idx:end])
			}

			// Find TOC
			idx = strings.Index(content, "TOC")
			if idx > 0 {
				start := idx - 200
				if start < 0 {
					start = 0
				}
				end := idx + 1000
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("=== Around TOC ===\n%s\n\n", content[start:end])
			}
		}
	}
}
