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

	// Check all slides for white text that should be dark
	// Focus on: transparent shapes below title area with white text
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

			// Split by shapes
			parts := strings.Split(content, "</p:sp>")
			for _, part := range parts {
				// Check if shape has noFill
				hasNoFill := strings.Contains(part, "<a:noFill/>")
				if !hasNoFill {
					continue
				}

				// Check if shape has white text
				if !strings.Contains(part, `val="FFFFFF"`) {
					continue
				}

				// Extract position
				offIdx := strings.Index(part, `<a:off x="`)
				if offIdx < 0 {
					continue
				}
				xEnd := strings.Index(part[offIdx+10:], `"`)
				if xEnd < 0 {
					continue
				}

				yIdx := strings.Index(part[offIdx:], `y="`)
				if yIdx < 0 {
					continue
				}
				yEnd := strings.Index(part[offIdx+yIdx+3:], `"`)
				if yEnd < 0 {
					continue
				}
				y := part[offIdx+yIdx+3 : offIdx+yIdx+3+yEnd]

				// Extract text
				texts := extractAllTexts(part)
				whiteTexts := []string{}
				for _, t := range texts {
					if t.color == "FFFFFF" && strings.TrimSpace(t.text) != "" {
						whiteTexts = append(whiteTexts, t.text)
					}
				}

				if len(whiteTexts) > 0 {
					allText := strings.Join(whiteTexts, " | ")
					if len(allText) > 100 {
						allText = allText[:100] + "..."
					}
					fmt.Printf("Slide %d: noFill shape at y=%s with WHITE text: %s\n", si, y, allText)
				}
			}
		}
	}
}

type textInfo struct {
	text  string
	color string
}

func extractAllTexts(shape string) []textInfo {
	var results []textInfo
	idx := 0
	for {
		ri := strings.Index(shape[idx:], "<a:r>")
		if ri < 0 {
			break
		}
		idx += ri
		re := strings.Index(shape[idx:], "</a:r>")
		if re < 0 {
			break
		}
		run := shape[idx : idx+re]

		color := ""
		ci := strings.Index(run, `srgbClr val="`)
		if ci >= 0 {
			ce := strings.Index(run[ci+13:], `"`)
			if ce >= 0 {
				color = run[ci+13 : ci+13+ce]
			}
		}

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

		if text != "" {
			results = append(results, textInfo{text: text, color: color})
		}
		idx += re + 6
	}
	return results
}
