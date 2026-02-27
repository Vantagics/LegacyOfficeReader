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
		if f.Name != "ppt/slides/slide26.xml" {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Split by shapes and check first 20
		parts := strings.Split(content, "</p:sp>")
		for i := 0; i < len(parts)-1 && i < 20; i++ {
			part := parts[i]
			hasNoFill := strings.Contains(part, "<a:noFill/>")
			hasSolidFill := strings.Contains(part, "<a:solidFill>")
			
			// Get fill color
			fillColor := "none"
			if hasSolidFill {
				ci := strings.Index(part, `<a:solidFill><a:srgbClr val="`)
				if ci >= 0 {
					ce := strings.Index(part[ci+28:], `"`)
					fillColor = part[ci+28 : ci+28+ce]
				}
			}
			if hasNoFill {
				fillColor = "noFill"
			}

			// Get text
			text := ""
			ti := strings.Index(part, "<a:t>")
			if ti >= 0 {
				te := strings.Index(part[ti+5:], "</a:t>")
				if te >= 0 {
					text = part[ti+5 : ti+5+te]
					if len(text) > 30 {
						text = text[:30]
					}
				}
			}

			fmt.Printf("Shape %d: fill=%s text=%q\n", i, fillColor, text)
		}
	}
}
