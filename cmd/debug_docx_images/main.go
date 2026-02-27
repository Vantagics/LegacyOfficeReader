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

			// Count inline and anchor images
			inlineCount := strings.Count(content, "<wp:inline")
			anchorCount := strings.Count(content, "<wp:anchor")
			fmt.Printf("Inline images: %d\n", inlineCount)
			fmt.Printf("Anchor images: %d\n", anchorCount)

			// Find all image references
			idx := 0
			imgCount := 0
			for {
				next := strings.Index(content[idx:], `r:embed="rImg`)
				if next < 0 {
					break
				}
				idx += next
				end := strings.Index(content[idx:], `"`)
				if end < 0 {
					break
				}
				end2 := strings.Index(content[idx+end+1:], `"`)
				relID := content[idx+len(`r:embed="`):idx+end+1+end2]
				fmt.Printf("  Image ref: %s\n", relID)
				imgCount++
				idx += end + 1
			}
			fmt.Printf("Total image references: %d\n", imgCount)

			// Find paragraphs with "部署拓扑图"
			targets := []string{"部署拓扑图", "未知威胁检测", "本地威胁发现", "文件威胁检测"}
			for _, target := range targets {
				tidx := strings.Index(content, target)
				if tidx >= 0 {
					pStart := strings.LastIndex(content[:tidx], "<w:p>")
					pEnd := strings.Index(content[tidx:], "</w:p>")
					if pStart >= 0 && pEnd >= 0 {
						para := content[pStart : tidx+pEnd+6]
						if len(para) > 400 {
							para = para[:400] + "..."
						}
						fmt.Printf("\n=== %s ===\n%s\n", target, para)
					}
				}
			}
		}
	}
}
