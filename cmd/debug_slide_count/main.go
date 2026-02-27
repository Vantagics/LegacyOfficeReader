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
	fmt.Printf("PPT slides: %d\n", len(slides))

	// Check if any slide has no shapes
	for i, s := range slides {
		shapes := s.GetShapes()
		if len(shapes) == 0 {
			fmt.Printf("  Slide %d: 0 shapes\n", i+1)
		}
	}
}
