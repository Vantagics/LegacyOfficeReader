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
	sl := slides[3] // slide 4
	for si, sh := range sl.GetShapes() {
		if sh.FillColor == "CFD5EA" && len(sh.Paragraphs) > 0 {
			fmt.Printf("Shape[%d] type=%d fill=%q noFill=%v fillRaw=0x%08X\n",
				si, sh.ShapeType, sh.FillColor, sh.NoFill, sh.FillColorRaw)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					t := run.Text
					if len(t) > 30 {
						t = t[:30] + "..."
					}
					fmt.Printf("  P[%d]R[%d] color=%q raw=0x%08X sz=%d: %q\n",
						pi, ri, run.Color, run.ColorRaw, run.FontSize, t)
				}
			}
		}
	}
	
	// Also check FFD966 fills
	fmt.Println("\n--- FFD966 fills ---")
	for si, sh := range sl.GetShapes() {
		if sh.FillColor == "FFD966" && len(sh.Paragraphs) > 0 {
			fmt.Printf("Shape[%d] type=%d fill=%q noFill=%v\n", si, sh.ShapeType, sh.FillColor, sh.NoFill)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					t := run.Text
					if len(t) > 30 {
						t = t[:30] + "..."
					}
					fmt.Printf("  P[%d]R[%d] color=%q raw=0x%08X: %q\n",
						pi, ri, run.Color, run.ColorRaw, t)
				}
			}
		}
	}
}
