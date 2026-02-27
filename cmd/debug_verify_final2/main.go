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

			// Count top-level paragraphs by walking the XML
			pCount := 0
			pos := 0
			anchorImgParas := 0
			inlineImgParas := 0
			textboxParas := 0
			pageBreakParas := 0
			headingParas := 0
			tocParas := 0
			emptyParas := 0

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

							if strings.Contains(pContent, "<wp:anchor") && !strings.Contains(pContent, "<w:txbxContent>") {
								anchorImgParas++
							}
							if strings.Contains(pContent, "<wp:inline") {
								inlineImgParas++
							}
							if strings.Contains(pContent, "<w:txbxContent>") {
								textboxParas++
							}
							if strings.Contains(pContent, `w:type="page"`) {
								pageBreakParas++
							}
							if strings.Contains(pContent, `w:val="Heading`) {
								headingParas++
							}
							if strings.Contains(pContent, `w:val="TOC`) {
								tocParas++
							}

							// Check if empty (no text, no images)
							hasText := false
							parts := strings.Split(pContent, "<w:t")
							for _, part := range parts[1:] {
								tStart := strings.Index(part, ">")
								tEnd := strings.Index(part, "</w:t>")
								if tStart >= 0 && tEnd > tStart {
									t := part[tStart+1 : tEnd]
									if strings.TrimSpace(t) != "" {
										hasText = true
										break
									}
								}
							}
							if !hasText && !strings.Contains(pContent, "<wp:") && !strings.Contains(pContent, `w:type="page"`) {
								emptyParas++
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

			fmt.Printf("Total top-level paragraphs: %d\n", pCount)
			fmt.Printf("Paragraphs with anchor images: %d\n", anchorImgParas)
			fmt.Printf("Paragraphs with inline images: %d\n", inlineImgParas)
			fmt.Printf("Paragraphs with textboxes: %d\n", textboxParas)
			fmt.Printf("Paragraphs with page breaks: %d\n", pageBreakParas)
			fmt.Printf("Heading paragraphs: %d\n", headingParas)
			fmt.Printf("TOC paragraphs: %d\n", tocParas)
			fmt.Printf("Empty paragraphs: %d\n", emptyParas)

			// Verify sectPr at end
			sectPrIdx := strings.LastIndex(content, "<w:sectPr>")
			if sectPrIdx > 0 {
				sectPr := content[sectPrIdx:]
				if len(sectPr) > 500 {
					sectPr = sectPr[:500]
				}
				fmt.Printf("\nFinal sectPr:\n%s\n", sectPr)
			}
		}
	}
}
