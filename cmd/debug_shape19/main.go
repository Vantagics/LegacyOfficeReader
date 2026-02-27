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
	shapes := sl.GetShapes()
	
	// Print ALL shapes with their details
	for si, sh := range shapes {
		fmt.Printf("Shape[%d] type=%d text=%v img=%v fill=%q noFill=%v paras=%d pos=(%d,%d) sz=(%d,%d)\n",
			si, sh.ShapeType, sh.IsText, sh.IsImage, sh.FillColor, sh.NoFill, len(sh.Paragraphs),
			sh.Left, sh.Top, sh.Width, sh.Height)
		for pi, para := range sh.Paragraphs {
			for ri, run := range para.Runs {
				t := run.Text
				if len(t) > 30 {
					t = t[:30] + "..."
				}
				fmt.Printf("  P[%d]R[%d] color=%q raw=0x%08X sz=%d font=%q: %q\n",
					pi, ri, run.Color, run.ColorRaw, run.FontSize, run.FontName, t)
			}
		}
	}
	
	// Now check what the PPTX shape id=21 corresponds to
	// PPTX assigns spID starting from 2, so id=21 = shape index 19
	fmt.Printf("\n--- Shape[19] detail ---\n")
	if len(shapes) > 19 {
		sh := shapes[19]
		fmt.Printf("type=%d text=%v img=%v fill=%q noFill=%v paras=%d\n",
			sh.ShapeType, sh.IsText, sh.IsImage, sh.FillColor, sh.NoFill, len(sh.Paragraphs))
	}
}
