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
	sl := slides[12]
	sh := sl.GetShapes()[2]
	fmt.Printf("Shape[2] type=%d text=%v paras=%d\n", sh.ShapeType, sh.IsText, len(sh.Paragraphs))
	for pi, para := range sh.Paragraphs {
		fmt.Printf("  P[%d] runs=%d\n", pi, len(para.Runs))
		for ri, run := range para.Runs {
			fmt.Printf("    R[%d] text=%q color=%q raw=0x%08X sz=%d\n",
				ri, run.Text, run.Color, run.ColorRaw, run.FontSize)
		}
	}

	// Also check shape 6 which has fill 8FAADC
	sh6 := sl.GetShapes()[6]
	fmt.Printf("\nShape[6] type=%d text=%v fill=%q paras=%d\n", sh6.ShapeType, sh6.IsText, sh6.FillColor, len(sh6.Paragraphs))
	for pi, para := range sh6.Paragraphs {
		fmt.Printf("  P[%d] runs=%d\n", pi, len(para.Runs))
		for ri, run := range para.Runs {
			fmt.Printf("    R[%d] text=%q color=%q raw=0x%08X sz=%d\n",
				ri, run.Text, run.Color, run.ColorRaw, run.FontSize)
		}
	}
}
