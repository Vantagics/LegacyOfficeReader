package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check specific slides for issues
	for _, slideNum := range []int{1, 2, 5, 10, 30, 50, 60, 70, 71} {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Extract all text content
				var texts []string
				for {
					idx := strings.Index(content, "<a:t>")
					if idx < 0 {
						idx = strings.Index(content, `<a:t xml:space="preserve">`)
						if idx < 0 {
							break
						}
						idx += len(`<a:t xml:space="preserve">`)
					} else {
						idx += len("<a:t>")
					}
					end := strings.Index(content[idx:], "</a:t>")
					if end < 0 {
						break
					}
					text := content[idx : idx+end]
					if len(text) > 60 {
						text = text[:60] + "..."
					}
					texts = append(texts, text)
					content = content[idx+end+len("</a:t>"):]
				}

				fmt.Printf("Slide %d: %d text runs\n", slideNum, len(texts))
				for i, t := range texts {
					if i > 5 {
						fmt.Printf("  ... and %d more\n", len(texts)-6)
						break
					}
					fmt.Printf("  [%d] %q\n", i, t)
				}
			}
		}
	}
}
