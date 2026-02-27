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

			// Find "修订记录"
			idx := strings.Index(content, "修订记录")
			if idx >= 0 {
				// Find the paragraph containing it
				pStart := strings.LastIndex(content[:idx], "<w:p>")
				if pStart < 0 {
					pStart = strings.LastIndex(content[:idx], "<w:p ")
				}
				pEnd := strings.Index(content[idx:], "</w:p>")
				if pEnd >= 0 {
					pEnd += idx + 6
				}
				if pStart >= 0 && pEnd > pStart {
					fmt.Printf("修订记录 paragraph:\n%s\n", content[pStart:pEnd])
				}
			} else {
				fmt.Println("修订记录 NOT FOUND")
			}

			// Find "状态：C-创建" (the legend text after the table)
			idx2 := strings.Index(content, "状态")
			if idx2 >= 0 {
				pStart := strings.LastIndex(content[:idx2], "<w:p>")
				if pStart < 0 {
					pStart = strings.LastIndex(content[:idx2], "<w:p ")
				}
				pEnd := strings.Index(content[idx2:], "</w:p>")
				if pEnd >= 0 {
					pEnd += idx2 + 6
				}
				if pStart >= 0 && pEnd > pStart {
					para := content[pStart:pEnd]
					if len(para) > 300 {
						para = para[:300] + "..."
					}
					fmt.Printf("\n状态 paragraph:\n%s\n", para)
				}
			}
		}
	}
}
