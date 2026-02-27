package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	f, _ := zip.OpenReader("testfie/test.pptx")
	defer f.Close()

	for _, zf := range f.File {
		if zf.Name == "ppt/slides/slide2.xml" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			xml := string(data)

			// Find all <p:sp> elements and dump them
			idx := 0
			spNum := 0
			for {
				start := strings.Index(xml[idx:], "<p:sp>")
				if start < 0 {
					break
				}
				start += idx
				end := strings.Index(xml[start:], "</p:sp>")
				if end < 0 {
					break
				}
				end += start + len("</p:sp>")
				snippet := xml[start:end]
				spNum++

				// Extract text
				textParts := []string{}
				tIdx := 0
				for {
					tStart := strings.Index(snippet[tIdx:], "<a:t>")
					if tStart < 0 {
						break
					}
					tStart += tIdx + 5
					tEnd := strings.Index(snippet[tStart:], "</a:t>")
					if tEnd < 0 {
						break
					}
					textParts = append(textParts, snippet[tStart:tStart+tEnd])
					tIdx = tStart + tEnd
				}
				text := strings.Join(textParts, " | ")

				// Only show text shapes with underline geometry
				if strings.Contains(text, "背景") || strings.Contains(text, "产品") || strings.Contains(text, "典型") || strings.Contains(text, "最佳") {
					fmt.Printf("=== Shape %d: '%s' ===\n", spNum, text)
					// Print the spPr section
					spPrStart := strings.Index(snippet, "<p:spPr>")
					spPrEnd := strings.Index(snippet, "</p:spPr>")
					if spPrStart >= 0 && spPrEnd >= 0 {
						spPr := snippet[spPrStart : spPrEnd+len("</p:spPr>")]
						fmt.Println(spPr)
					}
					fmt.Println()
				}

				idx = end
			}
		}
	}
}
