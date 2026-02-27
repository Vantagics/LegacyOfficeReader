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
	fmt.Printf("Total slides: %d\n", len(slides))

	if len(slides) < 26 {
		return
	}

	slide := slides[25]
	shapes := slide.GetShapes()
	fmt.Printf("Slide 26 raw shapes: %d\n", len(shapes))

	// Check how many have fill
	fillCount := 0
	noFillCount := 0
	for _, sh := range shapes {
		if sh.FillColor != "" && !sh.NoFill {
			fillCount++
		}
		if sh.NoFill {
			noFillCount++
		}
	}
	fmt.Printf("  With fill: %d, NoFill: %d, Other: %d\n", fillCount, noFillCount, len(shapes)-fillCount-noFillCount)

	// Check if shapes are being filtered
	// The issue might be in filterEmptySlides or similar
	// Let's check the formatted slide data
	masters := p.GetMasters()
	ref := slide.GetMasterRef()
	fmt.Printf("  MasterRef: %d\n", ref)
	if m, ok := masters[ref]; ok {
		fmt.Printf("  Master shapes: %d\n", len(m.Shapes))
		fmt.Printf("  Master bg: has=%v fill=%s imgIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
	}
}
