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
	slide := slides[7]
	shapes := slide.GetShapes()

	// Check all shapes, not just type=0
	for i, sh := range shapes {
		fmt.Printf("Shape %d: type=%d fill=%q noFill=%v isImage=%v imgIdx=%d isText=%v pos=(%d,%d) sz=(%d,%d)\n",
			i, sh.ShapeType, sh.FillColor, sh.NoFill, sh.IsImage, sh.ImageIdx, sh.IsText,
			sh.Left, sh.Top, sh.Width, sh.Height)
	}
}
