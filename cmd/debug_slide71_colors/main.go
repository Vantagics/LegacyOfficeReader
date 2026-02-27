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
	if len(slides) < 71 {
		fmt.Println("Not enough slides")
		return
	}

	slide := slides[70] // 0-indexed, slide 71
	fmt.Printf("Slide 71 - Master ref: %d\n", slide.GetMasterRef())
	fmt.Printf("Color scheme: %v\n", slide.GetColorScheme())

	shapes := slide.GetShapes()
	for si, sh := range shapes {
		if len(sh.Paragraphs) == 0 && !sh.IsImage {
			continue
		}
		text := ""
		if len(sh.Paragraphs) > 0 && len(sh.Paragraphs[0].Runs) > 0 {
			text = sh.Paragraphs[0].Runs[0].Text
			if len(text) > 30 {
				text = text[:30] + "..."
			}
		}
		fmt.Printf("Shape[%d]: fill=%s fillRaw=0x%08X line=%s lineRaw=0x%08X text=%q\n",
			si, sh.FillColor, sh.FillColorRaw, sh.LineColor, sh.LineColorRaw, text)
	}
}
