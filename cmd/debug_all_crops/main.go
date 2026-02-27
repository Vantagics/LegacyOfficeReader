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

	// Check master shapes
	masters := p.GetMasters()
	for ref, m := range masters {
		for i, sh := range m.Shapes {
			if sh.CropFromTop != 0 || sh.CropFromBottom != 0 || sh.CropFromLeft != 0 || sh.CropFromRight != 0 {
				fmt.Printf("Master %d shape[%d]: crop T=%d B=%d L=%d R=%d (imgIdx=%d)\n",
					ref, i, sh.CropFromTop, sh.CropFromBottom, sh.CropFromLeft, sh.CropFromRight, sh.ImageIdx)
			}
		}
	}

	// Check slide shapes
	slides := p.GetSlides()
	for si, s := range slides {
		shapes := s.GetShapes()
		for i, sh := range shapes {
			if sh.CropFromTop != 0 || sh.CropFromBottom != 0 || sh.CropFromLeft != 0 || sh.CropFromRight != 0 {
				fmt.Printf("Slide %d shape[%d]: crop T=%d B=%d L=%d R=%d (imgIdx=%d)\n",
					si+1, i, sh.CropFromTop, sh.CropFromBottom, sh.CropFromLeft, sh.CropFromRight, sh.ImageIdx)
			}
		}
	}
}
