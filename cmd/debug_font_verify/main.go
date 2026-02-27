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

	fonts := p.GetFonts()
	fmt.Printf("Font count: %d\n", len(fonts))
	for i, f := range fonts {
		fmt.Printf("  Font[%d]: ", i)
		for _, r := range f {
			fmt.Printf("U+%04X ", r)
		}
		fmt.Printf(" = %q\n", f)
	}

	// Check slide 1 shapes
	slides := p.GetSlides()
	if len(slides) > 0 {
		shapes := slides[0].GetShapes()
		for si, sh := range shapes {
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					fmt.Printf("Slide1 Shape[%d] Para[%d] Run[%d]: font=", si, pi, ri)
					for _, r := range run.FontName {
						fmt.Printf("U+%04X ", r)
					}
					fmt.Printf(" = %q size=%d color=%s\n", run.FontName, run.FontSize, run.Color)
				}
			}
		}
	}
}
