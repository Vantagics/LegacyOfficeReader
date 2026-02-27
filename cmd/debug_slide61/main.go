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

	for _, f := range zr.File {
		if f.Name != "ppt/slides/slide61.xml" {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Find all noFill shapes with FFFFFF text
		idx := 0
		for {
			pos := strings.Index(content[idx:], `val="FFFFFF"`)
			if pos < 0 {
				break
			}
			absPos := idx + pos

			// Look back for shape context
			spStart := strings.LastIndex(content[:absPos], "<p:sp>")
			if spStart < 0 {
				idx = absPos + 1
				continue
			}

			// Check if this is in a text run
			rprStart := strings.LastIndex(content[:absPos], "<a:rPr")
			if rprStart < 0 || rprStart < spStart {
				idx = absPos + 1
				continue
			}

			// Get shape fill
			spContent := content[spStart:absPos]
			hasNoFill := strings.Contains(spContent, "<a:noFill/>")

			if hasNoFill {
				// Get position
				offIdx := strings.Index(spContent, `<a:off x="`)
				xStr, yStr := "", ""
				if offIdx >= 0 {
					rest := spContent[offIdx+10:]
					xEnd := strings.Index(rest, `"`)
					if xEnd >= 0 {
						xStr = rest[:xEnd]
						yStart := strings.Index(rest, `y="`)
						if yStart >= 0 {
							yEnd := strings.Index(rest[yStart+3:], `"`)
							if yEnd >= 0 {
								yStr = rest[yStart+3 : yStart+3+yEnd]
							}
						}
					}
				}

				// Get text
				textStart := strings.Index(content[absPos:], "<a:t>")
				text := ""
				if textStart >= 0 {
					textEnd := strings.Index(content[absPos+textStart+5:], "</a:t>")
					if textEnd >= 0 {
						text = content[absPos+textStart+5 : absPos+textStart+5+textEnd]
					}
				}
				if len([]rune(text)) > 30 {
					text = string([]rune(text)[:30]) + "..."
				}

				fmt.Printf("WHITE noFill: x=%s y=%s text=%s\n", xStr, yStr, text)
			}

			idx = absPos + 1
		}
	}
}
