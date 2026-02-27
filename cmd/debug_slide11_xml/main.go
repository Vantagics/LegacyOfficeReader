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
		if f.Name != "ppt/slides/slide11.xml" {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Split by shapes
		parts := strings.Split(content, "</p:sp>")
		for i := 0; i < len(parts)-1; i++ {
			part := parts[i]
			
			// Get spPr fill
			spPrIdx := strings.Index(part, "<p:spPr>")
			if spPrIdx < 0 {
				continue
			}
			spPrEnd := strings.Index(part[spPrIdx:], "</p:spPr>")
			spPr := part[spPrIdx : spPrIdx+spPrEnd]
			
			lnIdx := strings.Index(spPr, "<a:ln")
			fillSection := spPr
			if lnIdx > 0 {
				fillSection = spPr[:lnIdx]
			}
			
			fillColor := "noFill"
			if strings.Contains(fillSection, "<a:solidFill>") {
				ci := strings.Index(fillSection, `srgbClr val="`)
				if ci >= 0 {
					ce := strings.Index(fillSection[ci+13:], `"`)
					fillColor = fillSection[ci+13 : ci+13+ce]
				}
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

			// Get text color
			textColor := ""
			tci := strings.Index(part, `<a:rPr`)
			if tci >= 0 {
				sci := strings.Index(part[tci:], `srgbClr val="`)
				if sci >= 0 {
					sce := strings.Index(part[tci+sci+13:], `"`)
					textColor = part[tci+sci+13 : tci+sci+13+sce]
				}
			}

			// Get position
			offIdx := strings.Index(part, `<a:off x="`)
			x, y := "?", "?"
			if offIdx >= 0 {
				xEnd := strings.Index(part[offIdx+10:], `"`)
				x = part[offIdx+10 : offIdx+10+xEnd]
				yIdx := strings.Index(part[offIdx:], `y="`)
				yEnd := strings.Index(part[offIdx+yIdx+3:], `"`)
				y = part[offIdx+yIdx+3 : offIdx+yIdx+3+yEnd]
			}

			fmt.Printf("Shape %d: fill=%s pos=(%s,%s) textColor=%s text=%q\n", i, fillColor, x, y, textColor, text)
		}

		// Also check pics
		picParts := strings.Split(content, "</p:pic>")
		for i := 0; i < len(picParts)-1; i++ {
			offIdx := strings.Index(picParts[i], `<a:off x="`)
			x, y := "?", "?"
			if offIdx >= 0 {
				xEnd := strings.Index(picParts[i][offIdx+10:], `"`)
				x = picParts[i][offIdx+10 : offIdx+10+xEnd]
				yIdx := strings.Index(picParts[i][offIdx:], `y="`)
				yEnd := strings.Index(picParts[i][offIdx+yIdx+3:], `"`)
				y = picParts[i][offIdx+yIdx+3 : offIdx+yIdx+3+yEnd]
			}
			fmt.Printf("PIC: pos=(%s,%s)\n", x, y)
		}
	}
}
