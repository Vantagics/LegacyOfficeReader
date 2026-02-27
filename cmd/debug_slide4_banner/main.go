package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	f, _ := zip.OpenReader("testfie/test.pptx")
	defer f.Close()

	for _, zf := range f.File {
		if zf.Name == "ppt/slides/slide4.xml" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			xml := string(data)

			// Find the yellow banner shape (fill FFD966)
			idx := 0
			spNum := 0
			for {
				start := strings.Index(xml[idx:], "<p:sp>")
				if start < 0 {
					break
				}
				start += idx
				end := strings.Index(xml[start:], "</p:sp>")
				if end < 0 {
					break
				}
				end += start + len("</p:sp>")
				snippet := xml[start:end]
				spNum++

				if strings.Contains(snippet, "FFD966") {
					fmt.Printf("=== Yellow banner (Shape %d) ===\n", spNum)
					fmt.Println(snippet[:min(len(snippet), 2000)])
					fmt.Println()
				}

				idx = end
			}
		}
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
