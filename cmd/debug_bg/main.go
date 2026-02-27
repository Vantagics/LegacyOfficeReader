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
	masters := p.GetMasters()

	// Check which masters are used and their backgrounds
	refCount := make(map[uint32]int)
	for _, s := range slides {
		refCount[s.GetMasterRef()]++
	}

	fmt.Println("=== Master backgrounds (used by slides) ===")
	for ref, count := range refCount {
		m, ok := masters[ref]
		if !ok {
			fmt.Printf("Master %d: NOT FOUND (used by %d slides)\n", ref, count)
			continue
		}
		fmt.Printf("Master %d (%d slides): bg=%v, fillColor=%s, imgIdx=%d, shapes=%d\n",
			ref, count, m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx, len(m.Shapes))
		
		// Check if any shape is a full-page image
		w, h := p.GetSlideSize()
		for i, sh := range m.Shapes {
			if sh.IsImage && sh.Width > int32(float64(w)*0.7) && sh.Height > int32(float64(h)*0.7) {
				fmt.Printf("  Shape %d: FULL-PAGE IMAGE (imgIdx=%d, %dx%d)\n", i, sh.ImageIdx, sh.Width, sh.Height)
			}
		}
	}

	// Check slide backgrounds
	fmt.Println("\n=== Slide backgrounds ===")
	hasBgCount := 0
	noBgCount := 0
	for _, s := range slides {
		bg := s.GetBackground()
		if bg.HasBackground {
			hasBgCount++
		} else {
			noBgCount++
		}
	}
	fmt.Printf("Slides with background: %d\n", hasBgCount)
	fmt.Printf("Slides without background (inherit from layout): %d\n", noBgCount)
}
