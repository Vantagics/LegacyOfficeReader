package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	// Map master refs to layout indices (same logic as pptconv)
	masterRefToLayoutIdx := make(map[uint32]int)
	layoutIdx := 0
	for _, s := range slides {
		ref := s.GetMasterRef()
		if _, ok := masterRefToLayoutIdx[ref]; !ok {
			masterRefToLayoutIdx[ref] = layoutIdx
			layoutIdx++
		}
	}

	// Print which slides use which layout
	layoutSlides := make(map[int][]int)
	for i, s := range slides {
		ref := s.GetMasterRef()
		li := masterRefToLayoutIdx[ref]
		layoutSlides[li] = append(layoutSlides[li], i+1)
	}

	for li := 0; li < layoutIdx; li++ {
		fmt.Printf("Layout %d: slides %v\n", li+1, layoutSlides[li])
	}
}
