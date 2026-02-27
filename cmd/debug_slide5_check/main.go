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

	// Check slide 5 PPTX output
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide5.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find all text runs with their colors
			idx := 0
			runNum := 0
			for {
				ri := strings.Index(content[idx:], "<a:r>")
				if ri < 0 {
					break
				}
				idx += ri
				re := strings.Index(content[idx:], "</a:r>")
				if re < 0 {
					break
				}
				run := content[idx : idx+re]

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

				if text != "" && len(text) > 80 {
					text = text[:80] + "..."
				}

				// Get font size
				sz := ""
				si := strings.Index(run, `sz="`)
				if si >= 0 {
					se := strings.Index(run[si+4:], `"`)
					if se >= 0 {
						sz = run[si+4 : si+4+se]
					}
				}

				if text != "" {
					fmt.Printf("Run[%d]: color=%s sz=%s text=%q\n", runNum, color, sz, text)
				}
				runNum++
				idx += re + 6
			}
		}
	}
}
