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

	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide71.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find all text runs with their colors
			parts := strings.Split(content, "<a:r>")
			for i, part := range parts {
				if i == 0 {
					continue
				}
				// Extract text
				tIdx := strings.Index(part, "<a:t>")
				tIdx2 := strings.Index(part, `<a:t xml:space="preserve">`)
				var text string
				if tIdx >= 0 {
					end := strings.Index(part[tIdx:], "</a:t>")
					if end >= 0 {
						text = part[tIdx+5 : tIdx+end]
					}
				} else if tIdx2 >= 0 {
					end := strings.Index(part[tIdx2:], "</a:t>")
					if end >= 0 {
						text = part[tIdx2+26 : tIdx2+end]
					}
				}

				// Extract color
				colorIdx := strings.Index(part, `<a:srgbClr val="`)
				color := ""
				if colorIdx >= 0 {
					color = part[colorIdx+16 : colorIdx+22]
				}

				// Extract font size
				szIdx := strings.Index(part, ` sz="`)
				sz := ""
				if szIdx >= 0 {
					end := strings.Index(part[szIdx+5:], `"`)
					if end >= 0 {
						sz = part[szIdx+5 : szIdx+5+end]
					}
				}

				if text != "" && len(text) < 60 {
					fmt.Printf("  text=%q color=%s sz=%s\n", text, color, sz)
				}
			}
		}
	}
}
