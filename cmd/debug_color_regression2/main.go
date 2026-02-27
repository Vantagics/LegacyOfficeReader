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

	// Check ALL slides for white text on near-white fills
	for _, zf := range f.File {
		if !strings.HasPrefix(zf.Name, "ppt/slides/slide") || !strings.HasSuffix(zf.Name, ".xml") {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		xml := string(data)

		issues := 0
		idx := 0
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

			// Get fill color from spPr
			spPrStart := strings.Index(snippet, "<p:spPr>")
			spPrEnd := strings.Index(snippet, "</p:spPr>")
			fillColor := ""
			if spPrStart >= 0 && spPrEnd >= 0 {
				spPr := snippet[spPrStart:spPrEnd]
				// First solidFill in spPr is the shape fill
				sfIdx := strings.Index(spPr, `<a:solidFill><a:srgbClr val="`)
				if sfIdx >= 0 {
					cStart := sfIdx + 28
					cEnd := strings.Index(spPr[cStart:], `"`)
					if cEnd >= 0 {
						fillColor = spPr[cStart : cStart+cEnd]
					}
				}
			}

			// Check if fill is near-white
			isNearWhite := fillColor == "FFFFFF" || fillColor == "E9EBF5" || fillColor == "CFD5EA" ||
				fillColor == "E7E6E6" || fillColor == "F2F2F2" || fillColor == "D9E2F3" ||
				fillColor == "DEEBF7" || fillColor == "FFF2CC"

			if isNearWhite {
				// Check for white text in txBody
				txBodyStart := strings.Index(snippet, "<p:txBody>")
				if txBodyStart >= 0 {
					txBody := snippet[txBodyStart:]
					if strings.Contains(txBody, `val="FFFFFF"`) {
						// Extract text
						textParts := []string{}
						tIdx := 0
						for {
							tStart := strings.Index(txBody[tIdx:], "<a:t>")
							if tStart < 0 {
								break
							}
							tStart += tIdx + 5
							tEnd := strings.Index(txBody[tStart:], "</a:t>")
							if tEnd < 0 {
								break
							}
							textParts = append(textParts, txBody[tStart:tStart+tEnd])
							tIdx = tStart + tEnd
						}
						text := strings.Join(textParts, " | ")
						if len(text) > 60 {
							text = text[:60] + "..."
						}
						if text != "" {
							issues++
							fmt.Printf("  %s: fill=%s WHITE text='%s'\n", zf.Name, fillColor, text)
						}
					}
				}
			}

			idx = end
		}
		if issues > 0 {
			fmt.Printf("  Total: %d issues\n\n", issues)
		}
	}
}
