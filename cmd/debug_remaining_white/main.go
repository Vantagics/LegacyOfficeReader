package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check specific high-white slides to understand what's still white
	for _, sn := range []int{4, 6, 9, 12, 15, 26, 30, 61, 71} {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", sn)
		for _, f := range zr.File {
			if f.Name != fname {
				continue
			}
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find all FFFFFF text runs and their context
			fmt.Printf("\n=== Slide %d ===\n", sn)
			idx := 0
			count := 0
			for count < 5 {
				// Find <a:solidFill><a:srgbClr val="FFFFFF"/>
				pos := strings.Index(content[idx:], `val="FFFFFF"`)
				if pos < 0 {
					break
				}
				absPos := idx + pos

				// Check if this is in a text run (near <a:rPr>) or a shape fill
				// Look backwards for context
				lookback := 200
				start := absPos - lookback
				if start < 0 { start = 0 }
				context := content[start:absPos]

				// Is this a text color (inside <a:rPr>) or a shape fill?
				isTextColor := strings.Contains(context, "<a:rPr") && !strings.Contains(context[strings.LastIndex(context, "<a:rPr"):], "</a:rPr>")

				if isTextColor {
					// Find the text
					textStart := strings.Index(content[absPos:], "<a:t>")
					textEnd := strings.Index(content[absPos:], "</a:t>")
					text := ""
					if textStart >= 0 && textEnd >= 0 && textEnd > textStart {
						text = content[absPos+textStart+5 : absPos+textEnd]
						if len([]rune(text)) > 40 {
							text = string([]rune(text)[:40]) + "..."
						}
					}

					// Find the shape fill (look for solidFill before this run)
					spStart := strings.LastIndex(content[:absPos], "<p:sp>")
					if spStart < 0 {
						spStart = strings.LastIndex(content[:absPos], "<p:pic>")
					}
					shapeFill := ""
					if spStart >= 0 {
						spContent := content[spStart:absPos]
						fillIdx := strings.Index(spContent, `<a:solidFill><a:srgbClr val="`)
						if fillIdx >= 0 {
							fillStart := fillIdx + len(`<a:solidFill><a:srgbClr val="`)
							fillEnd := strings.Index(spContent[fillStart:], `"`)
							if fillEnd >= 0 {
								shapeFill = spContent[fillStart : fillStart+fillEnd]
							}
						}
						if strings.Contains(spContent, "<a:noFill/>") && shapeFill == "" {
							shapeFill = "noFill"
						}
					}

					fmt.Printf("  WHITE text: fill=%s text=%s\n", shapeFill, text)
					count++
				}

				idx = absPos + 1
			}
		}
	}
}
