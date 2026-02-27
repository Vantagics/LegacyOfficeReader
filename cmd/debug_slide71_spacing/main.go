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
	slide := slides[70]
	shapes := slide.GetShapes()

	// Shape 36 - the large text box
	sh := shapes[36]
	fmt.Printf("Shape[36]: size=(%d,%d)\n", sh.Width, sh.Height)
	fmt.Printf("  TextMargins: L=%d T=%d R=%d B=%d\n",
		sh.TextMarginLeft, sh.TextMarginTop, sh.TextMarginRight, sh.TextMarginBottom)
	for pi, para := range sh.Paragraphs {
		fmt.Printf("  Para[%d]: lineSpacing=%d spaceBefore=%d spaceAfter=%d\n",
			pi, para.LineSpacing, para.SpaceBefore, para.SpaceAfter)
		for ri, run := range para.Runs {
			text := run.Text
			if len(text) > 50 {
				text = text[:50] + "..."
			}
			fmt.Printf("    Run[%d]: fontSize=%d text=%q\n", ri, run.FontSize, text)
		}
	}
}
