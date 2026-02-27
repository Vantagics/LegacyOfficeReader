package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	// Check specific slides
	for _, idx := range []int{0, 9, 14, 15, 20, 27, 39, 43, 45, 46, 52, 57, 60} {
		if idx >= len(slides) {
			continue
		}
		slide := slides[idx]
		shapes := slide.GetShapes()
		fmt.Printf("\n=== Slide %d (masterRef=%d, %d shapes) ===\n", idx+1, slide.GetMasterRef(), len(shapes))

		for i, sh := range shapes {
			hasText := len(sh.Paragraphs) > 0
			textPreview := ""
			firstColor := ""
			if hasText && len(sh.Paragraphs[0].Runs) > 0 {
				textPreview = sh.Paragraphs[0].Runs[0].Text
				firstColor = sh.Paragraphs[0].Runs[0].Color
				if len(textPreview) > 30 {
					textPreview = textPreview[:30] + "..."
				}
			}

			// Only show shapes with fill or text
			if sh.FillColor == "" && !hasText && !sh.IsImage {
				continue
			}

			fillInfo := "noColor"
			if sh.FillColor != "" && !sh.NoFill {
				fillInfo = fmt.Sprintf("fill=%s", sh.FillColor)
			} else if sh.NoFill {
				fillInfo = "noFill"
			}

			fmt.Printf("  %d: type=%d %s pos=(%d,%d) sz=(%d,%d)", i, sh.ShapeType, fillInfo, sh.Left, sh.Top, sh.Width, sh.Height)
			if sh.IsImage {
				fmt.Printf(" IMG=%d", sh.ImageIdx)
			}
			if hasText {
				fmt.Printf(" txtColor=%s text=%q", firstColor, textPreview)
			}
			fmt.Println()
		}
	}
}
