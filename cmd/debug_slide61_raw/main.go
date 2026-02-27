package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	pres, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := pres.GetSlides()
	if len(slides) < 61 {
		return
	}
	s := slides[60] // slide 61
	shapes := s.GetShapes()

	// Find shapes with white text and no fill near y=6067406
	for i, sh := range shapes {
		for _, p := range sh.Paragraphs {
			for _, r := range p.Runs {
				text := strings.TrimSpace(r.Text)
				if text == "" {
					continue
				}
				if r.Color == "FFFFFF" && sh.FillColor == "" {
					fmt.Printf("Shape[%d]: type=%d pos=(%d,%d) sz=(%d,%d) noFill=%v\n",
						i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.NoFill)
					fmt.Printf("  Text: %s  color=%s raw=0x%08X\n", text, r.Color, r.ColorRaw)
				}
			}
		}
	}
}
