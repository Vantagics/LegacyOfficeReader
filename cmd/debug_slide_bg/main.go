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
		fmt.Fprintf(os.Stderr, "Parse error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	tigerMaster := uint32(2147483734)
	for i, s := range slides {
		if s.GetMasterRef() != tigerMaster {
			continue
		}
		bg := s.GetBackground()
		if bg.HasBackground {
			fmt.Printf("Slide %d: has background, fillColor=%q, imageIdx=%d\n", i+1, bg.FillColor, bg.ImageIdx)
		}
	}
	fmt.Println("Done - slides with tiger master that have their own background")
}
