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

			// Find all <p:sp> elements
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

				// Extract text content
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

				hasCustGeom := strings.Contains(snippet, "custGeom")
				hasPrstGeom := strings.Contains(snippet, "prstGeom")
				hasNoFill := strings.Contains(snippet, "<a:noFill/>")
				
				// Extract line info
				lineInfo := "none"
				lineIdx := strings.Index(snippet, "<a:ln")
				if lineIdx >= 0 {
					lineEnd := strings.Index(snippet[lineIdx:], "</a:ln>")
					if lineEnd >= 0 {
						lineSnippet := snippet[lineIdx : lineIdx+lineEnd+len("</a:ln>")]
						if strings.Contains(lineSnippet, "noFill") {
							lineInfo = "noFill"
						} else if strings.Contains(lineSnippet, "solidFill") {
							// Extract color
							colorIdx := strings.Index(lineSnippet, `val="`)
							if colorIdx >= 0 {
								colorEnd := strings.Index(lineSnippet[colorIdx+5:], `"`)
								if colorEnd >= 0 {
									color := lineSnippet[colorIdx+5 : colorIdx+5+colorEnd]
									// Extract width
									wIdx := strings.Index(lineSnippet, `w="`)
									width := "default"
									if wIdx >= 0 {
										wEnd := strings.Index(lineSnippet[wIdx+3:], `"`)
										if wEnd >= 0 {
											width = lineSnippet[wIdx+3 : wIdx+3+wEnd]
										}
									}
									lineInfo = fmt.Sprintf("solid color=%s w=%s", color, width)
								}
							}
						}
					}
				}

				geom := "prstGeom"
				if hasCustGeom {
					geom = "custGeom"
				} else if hasPrstGeom {
					geom = "prstGeom"
				}

				fillInfo := "solid"
				if hasNoFill {
					fillInfo = "noFill"
				}

				fmt.Printf("Shape %d: geom=%s fill=%s line=%s text='%s'\n", spNum, geom, fillInfo, lineInfo, text)

				idx = end
			}
		}
	}
}
