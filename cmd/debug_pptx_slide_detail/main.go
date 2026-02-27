package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check specific slides for color issues
	colorRe := regexp.MustCompile(`srgbClr val="([A-F0-9]{6})"`)

	for si := 1; si <= 71; si++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Check for white text (FFFFFF) in solidFill within rPr (run properties)
				// Pattern: <a:rPr...><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
				whiteTextCount := 0
				idx := 0
				for {
					rprIdx := strings.Index(content[idx:], "<a:rPr")
					if rprIdx < 0 {
						break
					}
					idx += rprIdx
					rprEnd := strings.Index(content[idx:], "</a:rPr>")
					if rprEnd < 0 {
						break
					}
					rprContent := content[idx : idx+rprEnd]
					if strings.Contains(rprContent, `val="FFFFFF"`) && strings.Contains(rprContent, "solidFill") {
						whiteTextCount++
					}
					idx += rprEnd + 8
				}

				// Check for shapes with near-white fills that have white text
				// This is the main issue - white text on light fills
				hasIssue := false
				nearWhiteFills := []string{"E9EBF5", "CFD5EA", "F2F2F2", "E7E6E6", "EDEDED"}
				for _, fill := range nearWhiteFills {
					if strings.Contains(content, fmt.Sprintf(`val="%s"`, fill)) {
						// Check if there's white text nearby
						fillIdx := 0
						for {
							fi := strings.Index(content[fillIdx:], fmt.Sprintf(`val="%s"`, fill))
							if fi < 0 {
								break
							}
							fillIdx += fi + 10
							// Look ahead for white text in the same shape
							nextShape := strings.Index(content[fillIdx:], "</p:sp>")
							if nextShape < 0 {
								break
							}
							shapeContent := content[fillIdx : fillIdx+nextShape]
							if strings.Contains(shapeContent, `<a:solidFill><a:srgbClr val="FFFFFF"/>`) {
								hasIssue = true
								// Find the text
								textIdx := 0
								for {
									ti := strings.Index(shapeContent[textIdx:], "<a:t>")
									if ti < 0 {
										break
									}
									textIdx += ti + 5
									te := strings.Index(shapeContent[textIdx:], "</a:t>")
									if te < 0 {
										break
									}
									text := shapeContent[textIdx : textIdx+te]
									if len(text) > 60 {
										text = text[:60] + "..."
									}
									fmt.Printf("  Slide %d: WHITE text on %s fill: %q\n", si, fill, text)
									textIdx += te + 6
								}
							}
						}
					}
				}

				// Count all unique colors used in text
				allColors := colorRe.FindAllStringSubmatch(content, -1)
				colorCounts := make(map[string]int)
				for _, m := range allColors {
					colorCounts[m[1]]++
				}

				if hasIssue || whiteTextCount > 10 {
					fmt.Printf("Slide %d: whiteText=%d colors=%v\n", si, whiteTextCount, colorCounts)
				}
			}
		}
	}
}
