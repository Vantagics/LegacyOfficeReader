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
	if len(slides) < 26 {
		fmt.Println("Not enough slides")
		return
	}

	slide := slides[25] // 0-indexed, slide 26
	shapes := slide.GetShapes()

	// Check first few shapes with fill
	for i, sh := range shapes {
		if sh.FillColor != "" && len(sh.Paragraphs) > 0 {
			fmt.Printf("Shape %d: FillColor=%s NoFill=%v FillOpacity=%d\n", i, sh.FillColor, sh.NoFill, sh.FillOpacity)
			if i > 15 {
				break
			}
		}
	}
}
