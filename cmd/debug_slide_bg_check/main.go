package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	// Check specific slides for shape structure
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check slide 6 and slide 26 for background shapes
	for _, slideNum := range []int{6, 9, 26} {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		for _, f := range zr.File {
			if f.Name != name {
				continue
			}
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			fmt.Printf("\n=== Slide %d ===\n", slideNum)

			// Count shapes
			parts := strings.Split(content, "</p:sp>")
			fmt.Printf("Total shapes: %d\n", len(parts)-1)

			// Check for solidFill shapes (background rectangles)
			for i, part := range parts {
				if i >= len(parts)-1 {
					break
				}
				// Extract position
				offIdx := strings.Index(part, `<a:off x="`)
				if offIdx < 0 {
					continue
				}
				// Get x
				xEnd := strings.Index(part[offIdx+10:], `"`)
				x := part[offIdx+10 : offIdx+10+xEnd]
				// Get y
				yIdx := strings.Index(part[offIdx:], `y="`)
				yEnd := strings.Index(part[offIdx+yIdx+3:], `"`)
				y := part[offIdx+yIdx+3 : offIdx+yIdx+3+yEnd]
				// Get ext
				extIdx := strings.Index(part, `<a:ext cx="`)
				cx, cy := "?", "?"
				if extIdx >= 0 {
					cxEnd := strings.Index(part[extIdx+11:], `"`)
					cx = part[extIdx+11 : extIdx+11+cxEnd]
					cyIdx := strings.Index(part[extIdx:], `cy="`)
					cyEnd := strings.Index(part[extIdx+cyIdx+4:], `"`)
					cy = part[extIdx+cyIdx+4 : extIdx+cyIdx+4+cyEnd]
				}

				hasNoFill := strings.Contains(part, "<a:noFill/>")
				hasSolidFill := strings.Contains(part, "<a:solidFill>")
				hasWhiteText := strings.Contains(part, `val="FFFFFF"`)
				hasText := strings.Contains(part, "<a:t>") || strings.Contains(part, `<a:t xml:space=`)

				fillColor := ""
				if hasSolidFill {
					ci := strings.Index(part, `<a:solidFill><a:srgbClr val="`)
					if ci >= 0 {
						ce := strings.Index(part[ci+28:], `"`)
						fillColor = part[ci+28 : ci+28+ce]
					}
				}

				// Only show interesting shapes
				if hasSolidFill && !hasText && fillColor != "" {
					fmt.Printf("  Shape %d: FILL=%s pos=(%s,%s) sz=(%s,%s)\n", i, fillColor, x, y, cx, cy)
				}
				if hasNoFill && hasWhiteText && hasText {
					// Extract first text
					ti := strings.Index(part, "<a:t>")
					text := ""
					if ti >= 0 {
						te := strings.Index(part[ti+5:], "</a:t>")
						if te >= 0 {
							text = part[ti+5 : ti+5+te]
						}
					}
					if ti < 0 {
						ti = strings.Index(part, `<a:t xml:space="preserve">`)
						if ti >= 0 {
							te := strings.Index(part[ti+25:], "</a:t>")
							if te >= 0 {
								text = part[ti+25 : ti+25+te]
							}
						}
					}
					if len(text) > 30 {
						text = text[:30] + "..."
					}
					fmt.Printf("  Shape %d: noFill+WHITE pos=(%s,%s) sz=(%s,%s) text=%q\n", i, x, y, cx, cy, text)
				}
			}
		}
	}
}
