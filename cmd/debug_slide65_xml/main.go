package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "ppt/slides/slide65.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find text runs with their font sizes
			// Look for <a:r> elements
			parts := strings.Split(content, "<a:r>")
			for i, part := range parts {
				if i == 0 {
					continue
				}
				// Find the text
				tStart := strings.Index(part, "<a:t>")
				tEnd := strings.Index(part, "</a:t>")
				if tStart < 0 || tEnd < 0 {
					continue
				}
				text := part[tStart+5 : tEnd]
				if len(text) > 50 {
					text = text[:50] + "..."
				}

				// Find sz attribute
				szStr := ""
				szIdx := strings.Index(part, " sz=\"")
				if szIdx >= 0 && szIdx < tStart {
					szEnd := strings.Index(part[szIdx+5:], "\"")
					if szEnd >= 0 {
						szStr = part[szIdx+5 : szIdx+5+szEnd]
					}
				}

				// Find color
				colorStr := ""
				clrIdx := strings.Index(part, "<a:srgbClr val=\"")
				if clrIdx >= 0 && clrIdx < tStart {
					clrEnd := strings.Index(part[clrIdx+16:], "\"")
					if clrEnd >= 0 {
						colorStr = part[clrIdx+16 : clrIdx+16+clrEnd]
					}
				}

				fmt.Printf("Run[%d]: sz=%s color=%s text=%q\n", i, szStr, colorStr, text)
			}
			break
		}
	}
}
