package main

import (
	"fmt"
	"os"


	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	slides := p.GetSlides()

	// Check for potential text overflow issues
	for i, s := range slides {
		shapes := s.GetShapes()
		for j, sh := range shapes {
			if !sh.IsText || len(sh.Paragraphs) == 0 {
				continue
			}

			// Calculate total text length
			totalChars := 0
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					totalChars += len([]rune(run.Text))
				}
			}

			if totalChars == 0 {
				continue
			}

			// Check for very small shapes with lots of text
			heightPt := float64(sh.Height) / 12700.0 // EMU to points
			widthPt := float64(sh.Width) / 12700.0

			// Estimate if text might overflow
			// Rough estimate: CJK char needs ~12pt width, line height ~18pt
			maxFontSize := uint16(0)
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.FontSize > maxFontSize {
						maxFontSize = run.FontSize
					}
				}
			}

			if maxFontSize == 0 {
				// Check if this is a shape where font size estimation might be wrong
				if heightPt < 20 && totalChars > 5 {
					fmt.Printf("  Slide %d Shape %d: TINY shape (%.0f x %.0f pt) with %d chars, no font size\n",
						i+1, j, widthPt, heightPt, totalChars)
				}
			}

			// Check for shapes with very large text in small areas
			if heightPt > 0 && widthPt > 0 {
				fontSizePt := float64(maxFontSize) / 100.0
				if fontSizePt == 0 {
					fontSizePt = 14 // default estimate
				}
				charsPerLine := widthPt / (fontSizePt * 0.8)
				if charsPerLine < 1 {
					charsPerLine = 1
				}
				lines := float64(totalChars) / charsPerLine
				neededHeight := lines * fontSizePt * 1.3
				if neededHeight > heightPt*2 && totalChars > 20 {
					text := ""
					for _, para := range sh.Paragraphs {
						for _, run := range para.Runs {
							text += run.Text
						}
					}
					if len([]rune(text)) > 40 {
						text = string([]rune(text)[:40]) + "..."
					}
					fmt.Printf("  Slide %d Shape %d: potential overflow (%.0fx%.0f pt, %d chars, font=%.0f, need=%.0f) %q\n",
						i+1, j, widthPt, heightPt, totalChars, fontSizePt, neededHeight, text)
				}
			}
		}
	}
}
