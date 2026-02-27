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
	sl := slides[12] // slide 13
	fmt.Printf("Slide 13: %d shapes\n", len(sl.GetShapes()))
	for si, sh := range sl.GetShapes() {
		fmt.Printf("Shape[%d] type=%d text=%v img=%v fill=%q noFill=%v fillRaw=0x%08X pos=(%d,%d) sz=(%d,%d)\n",
			si, sh.ShapeType, sh.IsText, sh.IsImage, sh.FillColor, sh.NoFill, sh.FillColorRaw, sh.Left, sh.Top, sh.Width, sh.Height)
		if sh.IsText {
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					fmt.Printf("  P[%d]R[%d] color=%q raw=0x%08X sz=%d: %q\n",
						pi, ri, run.Color, run.ColorRaw, run.FontSize, run.Text)
				}
			}
		}
	}
}
