package main

import (
	"archive/zip"
	"fmt"
	"regexp"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		panic(err)
	}
	defer r.Close()

	// Check slides 32, 36, 46
	targets := []string{"ppt/slides/slide32.xml", "ppt/slides/slide36.xml", "ppt/slides/slide46.xml"}
	colorRe := regexp.MustCompile(`<a:solidFill><a:srgbClr val="([^"]+)"/></a:solidFill>`)

	for _, target := range targets {
		for _, f := range r.File {
			if f.Name == target {
				rc, _ := f.Open()
				buf := make([]byte, f.UncompressedSize64)
				rc.Read(buf)
				rc.Close()
				content := string(buf)

				fmt.Printf("\n=== %s ===\n", target)
				// Find text runs with their colors
				// Look for patterns like: <a:r><a:rPr ...><a:solidFill><a:srgbClr val="XXX"/></a:solidFill></a:rPr><a:t>TEXT</a:t></a:r>
				parts := strings.Split(content, "<a:r>")
				for _, part := range parts[1:] {
					endIdx := strings.Index(part, "</a:r>")
					if endIdx < 0 {
						continue
					}
					run := part[:endIdx]
					// Extract color
					colorMatch := colorRe.FindStringSubmatch(run)
					color := ""
					if colorMatch != nil {
						color = colorMatch[1]
					}
					// Extract text
					tStart := strings.Index(run, "<a:t>")
					tEnd := strings.Index(run, "</a:t>")
					text := ""
					if tStart >= 0 && tEnd > tStart {
						text = run[tStart+5 : tEnd]
					}
					if text == "" || text == " " {
						continue
					}
					// Only show colored runs (not default black)
					if color == "4472C4" || color == "ED7D31" || strings.Contains(text, "API传输") || strings.Contains(text, "数据流转监测") || strings.Contains(text, "数据安全管控") || strings.Contains(text, "条件规则") || strings.Contains(text, "基线规则") || strings.Contains(text, "流转数据监测") {
						if len(text) > 30 {
							text = text[:30]
						}
						fmt.Printf("  color=%s text=%q\n", color, text)
					}
				}
			}
		}
	}
}
