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
		fmt.Fprintf(os.Stderr, "Parse error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	s := slides[1] // slide 2
	shapes := s.GetShapes()
	fmt.Printf("Slide 2: %d shapes, masterRef=%d\n\n", len(shapes), s.GetMasterRef())

	for i, sh := range shapes {
		text := ""
		for _, p := range sh.Paragraphs {
			for _, r := range p.Runs {
				text += r.Text
			}
		}
		if len(text) > 60 {
			text = text[:60] + "..."
		}

		fmt.Printf("shape[%d]: type=%d pos=(%d,%d) size=(%d,%d)\n", i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height)
		fmt.Printf("  fill=%q noFill=%v line=%q noLine=%v lineW=%d\n", sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine, sh.LineWidth)
		if sh.IsImage {
			fmt.Printf("  IMAGE idx=%d\n", sh.ImageIdx)
		}
		if text != "" {
			fmt.Printf("  text=%q\n", text)
		}
		fmt.Println()
	}
}
