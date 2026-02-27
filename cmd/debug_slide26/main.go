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
	sl := slides[25] // slide 26
	for si, sh := range sl.GetShapes() {
		if (sh.FillColor == "A8DFFA" || sh.FillColor == "FFDC4C" || sh.FillColor == "FF974C" || sh.FillColor == "62D9AD" || sh.FillColor == "5AAEF3") && len(sh.Paragraphs) > 0 {
			t := ""
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if len(t) < 30 {
						t += run.Text
					}
				}
			}
			if len(t) > 30 {
				t = t[:30] + "..."
			}
			fmt.Printf("Shape[%d] fill=%q raw=0x%08X textColor=%q textRaw=0x%08X: %q\n",
				si, sh.FillColor, sh.FillColorRaw, sh.Paragraphs[0].Runs[0].Color, sh.Paragraphs[0].Runs[0].ColorRaw, t)
		}
	}
}
