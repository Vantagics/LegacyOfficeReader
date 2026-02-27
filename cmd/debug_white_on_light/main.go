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

	// Check slide 4 for white text on E9EBF5 fill
	checkSlides := []int{4, 9, 11, 12, 18, 26, 45, 63}
	for _, si := range checkSlides {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		for _, f := range zr.File {
			if f.Name != name {
				continue
			}
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Split into shapes
			shapes := strings.Split(content, "<p:sp>")
			for shIdx, shape := range shapes {
				if shIdx == 0 {
					continue
				}
				// Check if shape has a near-white fill
				nearWhiteFills := []string{"E9EBF5", "CFD5EA", "F2F2F2", "E7E6E6"}
				fillColor := ""
				for _, fill := range nearWhiteFills {
					if strings.Contains(shape, fmt.Sprintf(`val="%s"`, fill)) {
						fillColor = fill
						break
					}
				}
				if fillColor == "" {
					continue
				}

				// Check for white text in this shape
				if strings.Contains(shape, `<a:solidFill><a:srgbClr val="FFFFFF"/>`) {
					// Extract text
					texts := extractTexts(shape)
					for _, t := range texts {
						if t.color == "FFFFFF" {
							fmt.Printf("Slide %d Shape[%d]: fill=%s WHITE text: %q\n", si, shIdx, fillColor, truncate(t.text, 80))
						}
					}
				}
			}
		}
	}
}

type textInfo struct {
	text  string
	color string
}

func extractTexts(shape string) []textInfo {
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

		// Get color
		color := ""
		ci := strings.Index(run, `srgbClr val="`)
		if ci >= 0 {
			ce := strings.Index(run[ci+13:], `"`)
			if ce >= 0 {
				color = run[ci+13 : ci+13+ce]
			}
		}

		// Get text
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

func truncate(s string, n int) string {
	if len(s) > n {
		return s[:n] + "..."
	}
	return s
}
