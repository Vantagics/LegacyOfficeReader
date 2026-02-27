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

			// Count paragraphs (top-level only, not inside textboxes)
			// Simple approach: count <w:p> that are NOT inside <w:txbxContent>
			totalP := strings.Count(content, "<w:p>")
			totalPClose := strings.Count(content, "</w:p>")
			fmt.Printf("Total <w:p>: %d, </w:p>: %d\n", totalP, totalPClose)

			// Count anchored images
			anchorCount := strings.Count(content, "<wp:anchor")
			inlineCount := strings.Count(content, "<wp:inline")
			fmt.Printf("Anchored images: %d, Inline images: %d\n", anchorCount, inlineCount)

			// Count page breaks
			pageBreaks := strings.Count(content, `w:type="page"`)
			fmt.Printf("Page breaks: %d\n", pageBreaks)

			// Count section breaks
			sectPr := strings.Count(content, "<w:sectPr>")
			fmt.Printf("Section properties: %d\n", sectPr)

			// Check for wrapTopAndBottom
			wrapTB := strings.Count(content, "<wp:wrapTopAndBottom/>")
			fmt.Printf("wrapTopAndBottom: %d\n", wrapTB)

			// Check for wrapNone (textbox)
			wrapNone := strings.Count(content, "<wp:wrapNone/>")
			fmt.Printf("wrapNone: %d\n", wrapNone)

			// Check header/footer references
			hdrRef := strings.Count(content, "w:headerReference")
			ftrRef := strings.Count(content, "w:footerReference")
			fmt.Printf("Header refs: %d, Footer refs: %d\n", hdrRef, ftrRef)

			// Check TOC
			tocField := strings.Count(content, "TOC \\o")
			fmt.Printf("TOC fields: %d\n", tocField)

			// Show the first 15 top-level paragraphs (skip textbox inner paragraphs)
			// by finding <w:p> that are direct children of <w:body>
			bodyStart := strings.Index(content, "<w:body>")
			if bodyStart < 0 {
				return
			}
			body := content[bodyStart+8:]

			// Walk through body and extract top-level paragraphs
			pCount := 0
			pos := 0
			for pCount < 15 {
				pStart := strings.Index(body[pos:], "<w:p>")
				if pStart < 0 {
					break
				}
				pStart += pos

				// Find matching </w:p> (handle nesting)
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

							// Summarize paragraph
							hasAnchor := strings.Contains(pContent, "<wp:anchor")
							hasInline := strings.Contains(pContent, "<wp:inline")
							hasTxbx := strings.Contains(pContent, "<w:txbxContent>")
							hasPageBreak := strings.Contains(pContent, `w:type="page"`)

							// Extract text content (rough)
							text := ""
							parts := strings.Split(pContent, "<w:t")
							for _, part := range parts[1:] {
								tStart := strings.Index(part, ">")
								tEnd := strings.Index(part, "</w:t>")
								if tStart >= 0 && tEnd > tStart {
									t := part[tStart+1 : tEnd]
									// Skip textbox inner text
									if !strings.Contains(part[:tStart], "txbxContent") {
										text += t
									}
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
							if hasTxbx {
								flags += " TXBX"
							}
							if hasPageBreak {
								flags += " PB"
							}

							fmt.Printf("P[%2d]: text=%q%s\n", pCount+1, text, flags)
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
