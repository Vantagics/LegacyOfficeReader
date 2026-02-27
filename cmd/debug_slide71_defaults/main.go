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
	slide := slides[70] // slide 71
	defaults := slide.GetDefaultTextStyles()
	fmt.Printf("Slide 71 default text styles:\n")
	for i, d := range defaults {
		fmt.Printf("  Level %d: fontSize=%d fontName=%q bold=%v color=%s\n",
			i, d.FontSize, d.FontName, d.Bold, d.Color)
	}

	// Check shapes with fontSize=0
	shapes := slide.GetShapes()
	for si, sh := range shapes {
		for pi, para := range sh.Paragraphs {
			for ri, run := range para.Runs {
				if run.FontSize == 0 {
					text := run.Text
					if len(text) > 40 {
						text = text[:40] + "..."
					}
					fmt.Printf("Shape[%d] P%d/R%d: fontSize=0 text=%q shapeSize=(%d,%d)\n",
						si, pi, ri, text, sh.Width, sh.Height)
				}
			}
		}
	}
}
