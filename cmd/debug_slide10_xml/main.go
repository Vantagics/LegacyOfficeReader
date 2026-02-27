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
		if f.Name != "ppt/slides/slide15.xml" {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		fmt.Printf("Total sp: %d, pic: %d\n", strings.Count(content, "</p:sp>"), strings.Count(content, "</p:pic>"))

		// Show all shapes in order
		// Find all <p:sp> and <p:pic> elements with their positions in the string
		idx := 0
		shapeNum := 0
		for idx < len(content) {
			spIdx := strings.Index(content[idx:], "<p:sp>")
			picIdx := strings.Index(content[idx:], "<p:pic>")

			if spIdx < 0 && picIdx < 0 {
				break
			}

			if picIdx >= 0 && (spIdx < 0 || picIdx < spIdx) {
				// pic comes first
				start := idx + picIdx
				end := strings.Index(content[start:], "</p:pic>")
				if end < 0 {
					break
				}
				part := content[start : start+end]
				offIdx := strings.Index(part, `<a:off x="`)
				x, y := "?", "?"
				if offIdx >= 0 {
					xEnd := strings.Index(part[offIdx+10:], `"`)
					x = part[offIdx+10 : offIdx+10+xEnd]
					yIdx := strings.Index(part[offIdx:], `y="`)
					yEnd := strings.Index(part[offIdx+yIdx+3:], `"`)
					y = part[offIdx+yIdx+3 : offIdx+yIdx+3+yEnd]
				}
				fmt.Printf("  [%d] PIC pos=(%s,%s)\n", shapeNum, x, y)
				idx = start + end + 8
				shapeNum++
			} else {
				// sp comes first
				start := idx + spIdx
				end := strings.Index(content[start:], "</p:sp>")
				if end < 0 {
					break
				}
				part := content[start : start+end]

				// Get fill
				spPrIdx := strings.Index(part, "<p:spPr>")
				fillColor := "?"
				if spPrIdx >= 0 {
					spPrEnd := strings.Index(part[spPrIdx:], "</p:spPr>")
					spPr := part[spPrIdx : spPrIdx+spPrEnd]
					lnIdx := strings.Index(spPr, "<a:ln")
					fillSection := spPr
					if lnIdx > 0 {
						fillSection = spPr[:lnIdx]
					}
					if strings.Contains(fillSection, "<a:noFill/>") {
						fillColor = "noFill"
					} else if strings.Contains(fillSection, "<a:solidFill>") {
						ci := strings.Index(fillSection, `srgbClr val="`)
						if ci >= 0 {
							ce := strings.Index(fillSection[ci+13:], `"`)
							fillColor = fillSection[ci+13 : ci+13+ce]
						}
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

				offIdx := strings.Index(part, `<a:off x="`)
				x, y := "?", "?"
				if offIdx >= 0 {
					xEnd := strings.Index(part[offIdx+10:], `"`)
					x = part[offIdx+10 : offIdx+10+xEnd]
					yIdx := strings.Index(part[offIdx:], `y="`)
					yEnd := strings.Index(part[offIdx+yIdx+3:], `"`)
					y = part[offIdx+yIdx+3 : offIdx+yIdx+3+yEnd]
				}

				fmt.Printf("  [%d] SP fill=%s pos=(%s,%s) text=%q\n", shapeNum, fillColor, x, y, text)
				idx = start + end + 7
				shapeNum++
			}
		}
	}
}
