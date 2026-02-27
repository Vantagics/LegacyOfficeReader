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
	sl := slides[40] // slide 41
	for si, sh := range sl.GetShapes() {
		if sh.FillColor == "000000" && len(sh.Paragraphs) > 0 {
			t := ""
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t += run.Text
				}
			}
			fmt.Printf("Shape[%d] fill=%q noFill=%v opacity=%d lineColor=%q lineRaw=0x%08X: %q\n",
				si, sh.FillColor, sh.NoFill, sh.FillOpacity, sh.LineColor, sh.LineColorRaw, t)
		}
	}
}
