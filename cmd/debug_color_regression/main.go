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

	// Check a few slides for color issues
	// Slide 1 (cover), slide 4, slide 8, slide 10
	targets := []string{"slide1.xml", "slide4.xml", "slide8.xml", "slide10.xml", "slide2.xml"}
	for _, zf := range f.File {
		for _, target := range targets {
			if zf.Name == "ppt/slides/"+target {
				rc, _ := zf.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				xml := string(data)

				// Count white text on various fills
				whiteOnLight := 0
				whiteOnDark := 0
				blackOnLight := 0
				
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

					// Get fill color
					fillColor := ""
					fillIdx := strings.Index(snippet, `<a:solidFill><a:srgbClr val="`)
					if fillIdx >= 0 {
						fillStart := fillIdx + 28
						fillEnd := strings.Index(snippet[fillStart:], `"`)
						if fillEnd >= 0 {
							fillColor = snippet[fillStart : fillStart+fillEnd]
						}
					}

					// Check text colors
					hasWhiteText := false
					hasBlackText := false
					tIdx := 0
					for {
						rprStart := strings.Index(snippet[tIdx:], `<a:rPr `)
						if rprStart < 0 {
							break
						}
						rprStart += tIdx
						rprEnd := strings.Index(snippet[rprStart:], `</a:rPr>`)
						if rprEnd < 0 {
							rprEnd = strings.Index(snippet[rprStart:], `/>`)
							if rprEnd < 0 {
								break
							}
						}
						rprSnippet := snippet[rprStart : rprStart+rprEnd+2]
						if strings.Contains(rprSnippet, `val="FFFFFF"`) {
							hasWhiteText = true
						}
						if strings.Contains(rprSnippet, `val="000000"`) {
							hasBlackText = true
						}
						tIdx = rprStart + rprEnd + 2
					}

					isLightFill := fillColor != "" && (fillColor == "FFFFFF" || fillColor == "E9EBF5" || fillColor == "CFD5EA" || fillColor == "E7E6E6")
					isDarkFill := fillColor != "" && !isLightFill

					if hasWhiteText && isLightFill {
						whiteOnLight++
					}
					if hasWhiteText && isDarkFill {
						whiteOnDark++
					}
					if hasBlackText && isLightFill {
						blackOnLight++
					}

					idx = end
				}

				fmt.Printf("%s: whiteOnLight=%d whiteOnDark=%d blackOnLight=%d\n",
					target, whiteOnLight, whiteOnDark, blackOnLight)
			}
		}
	}
}
