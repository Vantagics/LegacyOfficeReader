package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	path := "testfie/test.docx"
	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find paragraphs containing specific text
			targets := []string{"修订记录", "目", "版权声明", "引言", "产品概述", "产品组成"}
			for _, target := range targets {
				idx := strings.Index(content, target)
				if idx < 0 {
					continue
				}
				// Find the enclosing <w:p>...</w:p>
				pStart := strings.LastIndex(content[:idx], "<w:p>")
				if pStart < 0 {
					pStart = strings.LastIndex(content[:idx], "<w:p ")
				}
				pEnd := strings.Index(content[idx:], "</w:p>")
				if pStart >= 0 && pEnd >= 0 {
					para := content[pStart : idx+pEnd+6]
					if len(para) > 800 {
						para = para[:800] + "..."
					}
					fmt.Printf("=== %s ===\n%s\n\n", target, para)
				}
			}
		}
	}
}
