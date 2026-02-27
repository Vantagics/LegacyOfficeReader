package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer r.Close()

	// Check numbering.xml
	for _, f := range r.File {
		if f.Name == "word/numbering.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("numbering.xml: %d bytes\n", len(data))
			content := string(data)
			// Count abstractNum and num
			absCount := strings.Count(content, "<w:abstractNum")
			numCount := strings.Count(content, "<w:num ")
			fmt.Printf("  abstractNum: %d, num: %d\n", absCount, numCount)
			// Show bullet definition
			if strings.Contains(content, "bullet") {
				fmt.Println("  Has bullet numbering definition")
			}
			if len(content) < 5000 {
				fmt.Println(content)
			}
		}
	}

	// Check document.xml for numPr
	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			numPrCount := strings.Count(content, "<w:numPr>")
			fmt.Printf("\ndocument.xml: <w:numPr> count = %d\n", numPrCount)

			// Find paragraphs with numPr and show their text
			bodyStart := strings.Index(content, "<w:body>")
			if bodyStart < 0 {
				return
			}
			body := content[bodyStart+8:]

			pCount := 0
			pos := 0
			for pos < len(body) {
				pStart := strings.Index(body[pos:], "<w:p>")
				if pStart < 0 {
					break
				}
				pStart += pos

				depth := 1
				searchPos := pStart + 5
				for depth > 0 && searchPos < len(body) {
					nextOpen := strings.Index(body[searchPos:], "<w:p>")
					nextClose := strings.Index(body[searchPos:], "</w:p>")
					if nextClose < 0 {
						break
					}
					if nextOpen >= 0 && nextOpen < nextClose {
						depth++
						searchPos += nextOpen + 5
					} else {
						depth--
						if depth == 0 {
							pEnd := searchPos + nextClose + 6
							pContent := body[pStart:pEnd]
							pCount++

							if strings.Contains(pContent, "<w:numPr>") {
								// Extract text
								text := ""
								parts := strings.Split(pContent, "<w:t")
								for _, part := range parts[1:] {
									tStart := strings.Index(part, ">")
									tEnd := strings.Index(part, "</w:t>")
									if tStart >= 0 && tEnd > tStart {
										text += part[tStart+1 : tEnd]
									}
								}
								if len(text) > 60 {
									text = text[:60] + "..."
								}

								// Extract numId and ilvl
								numIdStart := strings.Index(pContent, `<w:numId w:val="`)
								numId := ""
								if numIdStart >= 0 {
									numIdStart += 16
									numIdEnd := strings.Index(pContent[numIdStart:], `"`)
									if numIdEnd >= 0 {
										numId = pContent[numIdStart : numIdStart+numIdEnd]
									}
								}
								ilvlStart := strings.Index(pContent, `<w:ilvl w:val="`)
								ilvl := ""
								if ilvlStart >= 0 {
									ilvlStart += 15
									ilvlEnd := strings.Index(pContent[ilvlStart:], `"`)
									if ilvlEnd >= 0 {
										ilvl = pContent[ilvlStart : ilvlStart+ilvlEnd]
									}
								}

								hasHeading := strings.Contains(pContent, "Heading")
								hFlag := ""
								if hasHeading {
									hFlag = " HEADING"
								}

								fmt.Printf("  P[%3d] numId=%s ilvl=%s%s text=%q\n", pCount, numId, ilvl, hFlag, text)
							}

							pos = pEnd
						} else {
							searchPos += nextClose + 6
						}
					}
				}
				if depth > 0 {
					break
				}
			}
		}
	}
}
