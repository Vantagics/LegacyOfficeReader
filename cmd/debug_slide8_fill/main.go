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

	for _, zf := range r.File {
		if zf.Name != "ppt/slides/slide8.xml" {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Find all <p:sp> elements and extract their fill info
		parts := strings.Split(content, "<p:sp>")
		for i, part := range parts {
			if i == 0 {
				continue // skip preamble
			}
			// Get shape name
			nameIdx := strings.Index(part, `name="`)
			name := ""
			if nameIdx >= 0 {
				nameEnd := strings.Index(part[nameIdx+6:], `"`)
				if nameEnd >= 0 {
					name = part[nameIdx+6 : nameIdx+6+nameEnd]
				}
			}

			// Check for custGeom
			hasCustGeom := strings.Contains(part, "custGeom")

			// Find fill type
			fillType := "unknown"
			if strings.Contains(part, "<a:noFill/>") {
				// Check if it's in spPr (not in ln)
				spPrIdx := strings.Index(part, "<p:spPr>")
				noFillIdx := strings.Index(part, "<a:noFill/>")
				lnIdx := strings.Index(part, "<a:ln>")
				if spPrIdx >= 0 && noFillIdx > spPrIdx {
					if lnIdx < 0 || noFillIdx < lnIdx {
						fillType = "noFill"
					}
				}
			}
			if strings.Contains(part, "<a:solidFill>") {
				solidIdx := strings.Index(part, "<a:solidFill>")
				lnIdx := strings.Index(part, "<a:ln>")
				if lnIdx < 0 || solidIdx < lnIdx {
					// Extract color
					colorIdx := strings.Index(part[solidIdx:], `val="`)
					if colorIdx >= 0 {
						colorEnd := strings.Index(part[solidIdx+colorIdx+5:], `"`)
						if colorEnd >= 0 {
							fillType = "solidFill:" + part[solidIdx+colorIdx+5:solidIdx+colorIdx+5+colorEnd]
						}
					}
				}
			}

			if hasCustGeom {
				fmt.Printf("Shape: %-20s custGeom=true  fill=%s\n", name, fillType)
			}
		}
	}
}
