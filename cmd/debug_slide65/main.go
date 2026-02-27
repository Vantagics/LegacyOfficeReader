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
	if len(slides) < 65 {
		fmt.Println("Not enough slides")
		return
	}

	slide := slides[64] // 0-indexed
	shapes := slide.GetShapes()

	fmt.Printf("Slide 65: %d shapes\n\n", len(shapes))
	for i, s := range shapes {
		hasText := false
		for _, p := range s.Paragraphs {
			for _, r := range p.Runs {
				if len(r.Text) > 0 {
					hasText = true
					break
				}
			}
			if hasText {
				break
			}
		}
		if !hasText {
			continue
		}

		fmt.Printf("Shape[%d]: type=%d pos=(%d,%d) size=(%d,%d)\n", i, s.ShapeType, s.Left, s.Top, s.Width, s.Height)
		fmt.Printf("  FillColor=%q NoFill=%v IsImage=%v\n", s.FillColor, s.NoFill, s.IsImage)
		for pi, para := range s.Paragraphs {
			for ri, run := range para.Runs {
				if len(run.Text) == 0 {
					continue
				}
				text := run.Text
				if len(text) > 40 {
					text = text[:40] + "..."
				}
				fmt.Printf("  P[%d]R[%d]: sz=%d color=%q bold=%v text=%q\n",
					pi, ri, run.FontSize, run.Color, run.Bold, text)
			}
		}
		fmt.Println()
	}
}
