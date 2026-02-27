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
	if len(slides) < 26 {
		fmt.Println("Not enough slides")
		return
	}

	slide := slides[25] // 0-indexed, slide 26
	shapes := slide.GetShapes()
	fmt.Printf("Slide 26: %d shapes\n", len(shapes))

	for i, sh := range shapes {
		hasText := len(sh.Paragraphs) > 0
		textPreview := ""
		if hasText && len(sh.Paragraphs) > 0 && len(sh.Paragraphs[0].Runs) > 0 {
			textPreview = sh.Paragraphs[0].Runs[0].Text
			if len(textPreview) > 30 {
				textPreview = textPreview[:30] + "..."
			}
		}

		fillInfo := "noFill"
		if sh.FillColor != "" && !sh.NoFill {
			fillInfo = fmt.Sprintf("fill=%s", sh.FillColor)
		} else if sh.NoFill {
			fillInfo = "noFill"
		} else if sh.FillColor == "" {
			fillInfo = "noColor"
		}

		if sh.FillColor != "" || hasText {
			fmt.Printf("  Shape %d: type=%d %s pos=(%d,%d) sz=(%d,%d)", i, sh.ShapeType, fillInfo, sh.Left, sh.Top, sh.Width, sh.Height)
			if hasText {
				// Check first run color
				if len(sh.Paragraphs[0].Runs) > 0 {
					r := sh.Paragraphs[0].Runs[0]
					fmt.Printf(" color=%s raw=0x%08X", r.Color, r.ColorRaw)
				}
				fmt.Printf(" text=%q", textPreview)
			}
			fmt.Println()
		}
	}
}
