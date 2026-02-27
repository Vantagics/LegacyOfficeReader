package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	slide := slides[10] // slide 11
	shapes := slide.GetShapes()
	fmt.Printf("Slide 11: %d shapes, masterRef=%d\n", len(shapes), slide.GetMasterRef())

	for i, sh := range shapes {
		hasText := len(sh.Paragraphs) > 0
		textPreview := ""
		firstColor := ""
		firstColorRaw := uint32(0)
		if hasText && len(sh.Paragraphs[0].Runs) > 0 {
			textPreview = sh.Paragraphs[0].Runs[0].Text
			firstColor = sh.Paragraphs[0].Runs[0].Color
			firstColorRaw = sh.Paragraphs[0].Runs[0].ColorRaw
			if len(textPreview) > 40 {
				textPreview = textPreview[:40] + "..."
			}
		}

		fillInfo := fmt.Sprintf("fill=%s noFill=%v", sh.FillColor, sh.NoFill)

		fmt.Printf("  Shape %d: type=%d %s pos=(%d,%d) sz=(%d,%d)", i, sh.ShapeType, fillInfo, sh.Left, sh.Top, sh.Width, sh.Height)
		if hasText {
			fmt.Printf(" color=%s raw=0x%08X text=%q", firstColor, firstColorRaw, textPreview)
		}
		if sh.IsImage {
			fmt.Printf(" IMG=%d", sh.ImageIdx)
		}
		fmt.Println()
	}
}
