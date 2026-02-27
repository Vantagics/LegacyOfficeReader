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

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			bodyStart := strings.Index(content, "<w:body>")
			if bodyStart < 0 {
				return
			}
			body := content[bodyStart+8:]

			// Walk through body and find paragraphs with images
			pCount := 0
			pos := 0
			for pos < len(body) {
				pStart := strings.Index(body[pos:], "<w:p>")
				if pStart < 0 {
					break
				}
				pStart += pos

				// Find matching </w:p>
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

							hasAnchor := strings.Contains(pContent, "<wp:anchor")
							hasInline := strings.Contains(pContent, "<wp:inline")
							hasPageBreak := strings.Contains(pContent, `w:type="page"`)
							hasSectPr := strings.Contains(pContent, "<w:sectPr>")

							if hasAnchor || hasInline || hasPageBreak || hasSectPr {
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

								flags := ""
								if hasAnchor {
									flags += " ANCHOR"
								}
								if hasInline {
									flags += " INLINE"
								}
								if hasPageBreak {
									flags += " PB"
								}
								if hasSectPr {
									flags += " SECT"
								}

								// Get image embed IDs
								embeds := []string{}
								idx := 0
								for {
									eStart := strings.Index(pContent[idx:], `r:embed="`)
									if eStart < 0 {
										break
									}
									eStart += idx + 9
									eEnd := strings.Index(pContent[eStart:], `"`)
									if eEnd >= 0 {
										embeds = append(embeds, pContent[eStart:eStart+eEnd])
									}
									idx = eStart + eEnd + 1
								}

								fmt.Printf("P[%3d]: text=%q embeds=%v%s\n", pCount+1, text, embeds, flags)
							}

							pCount++
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
