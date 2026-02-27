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

	totalIssues := 0
	for si := 1; si <= 71; si++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		for _, f := range zr.File {
			if f.Name != name {
				continue
			}
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Split by shapes (p:sp only, not p:pic)
			parts := strings.Split(content, "</p:sp>")
			for _, part := range parts {
				// Find the spPr section to check fill
				spPrIdx := strings.Index(part, "<p:spPr>")
				if spPrIdx < 0 {
					continue
				}
				spPrEnd := strings.Index(part[spPrIdx:], "</p:spPr>")
				if spPrEnd < 0 {
					continue
				}
				spPr := part[spPrIdx : spPrIdx+spPrEnd]

				// Check if shape fill (not line fill) is noFill
				// The shape fill comes before <a:ln> in spPr
				lnIdx := strings.Index(spPr, "<a:ln")
				fillSection := spPr
				if lnIdx > 0 {
					fillSection = spPr[:lnIdx]
				}

				shapeHasNoFill := strings.Contains(fillSection, "<a:noFill/>")
				if !shapeHasNoFill {
					continue // shape has a fill, text should be visible
				}

				// Check if shape has white text
				txBodyIdx := strings.Index(part, "<p:txBody>")
				if txBodyIdx < 0 {
					continue
				}
				txBody := part[txBodyIdx:]

				if !strings.Contains(txBody, `val="FFFFFF"`) {
					continue
				}

				// Extract text
				texts := extractWhiteTexts(txBody)
				if len(texts) == 0 {
					continue
				}

				// Extract position
				offIdx := strings.Index(part, `<a:off x="`)
				y := "?"
				if offIdx >= 0 {
					yIdx := strings.Index(part[offIdx:], `y="`)
					if yIdx >= 0 {
						yEnd := strings.Index(part[offIdx+yIdx+3:], `"`)
						y = part[offIdx+yIdx+3 : offIdx+yIdx+3+yEnd]
					}
				}

				allText := strings.Join(texts, " | ")
				if len(allText) > 80 {
					allText = allText[:80] + "..."
				}
				fmt.Printf("Slide %d: TRUE noFill+WHITE at y=%s: %s\n", si, y, allText)
				totalIssues++
			}
		}
	}
	fmt.Printf("\nTotal real issues: %d\n", totalIssues)
}

func extractWhiteTexts(txBody string) []string {
	var results []string
	idx := 0
	for {
		ri := strings.Index(txBody[idx:], "<a:r>")
		if ri < 0 {
			break
		}
		idx += ri
		re := strings.Index(txBody[idx:], "</a:r>")
		if re < 0 {
			break
		}
		run := txBody[idx : idx+re]

		// Check if this run has white color
		if !strings.Contains(run, `val="FFFFFF"`) {
			idx += re + 6
			continue
		}

		// Extract text
		text := ""
		ti := strings.Index(run, "<a:t>")
		if ti >= 0 {
			te := strings.Index(run[ti+5:], "</a:t>")
			if te >= 0 {
				text = run[ti+5 : ti+5+te]
			}
		}
		if ti < 0 {
			ti = strings.Index(run, `<a:t xml:space="preserve">`)
			if ti >= 0 {
				te := strings.Index(run[ti+25:], "</a:t>")
				if te >= 0 {
					text = run[ti+25 : ti+25+te]
				}
			}
		}

		if strings.TrimSpace(text) != "" {
			results = append(results, text)
		}
		idx += re + 6
	}
	return results
}
