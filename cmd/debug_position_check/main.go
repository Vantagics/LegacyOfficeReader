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
	w, h := p.GetSlideSize()

	fmt.Printf("Slide size: %d x %d EMU\n", w, h)

	// Check for shapes with negative coordinates or outside slide area
	for i, s := range slides {
		shapes := s.GetShapes()
		for si, sh := range shapes {
			if sh.Left < 0 || sh.Top < 0 {
				fmt.Printf("Slide %d Shape %d: NEGATIVE pos (%d,%d) size (%d,%d) type=%d\n",
					i+1, si, sh.Left, sh.Top, sh.Width, sh.Height, sh.ShapeType)
			}
			if sh.Left+sh.Width > int32(w)*2 || sh.Top+sh.Height > int32(h)*2 {
				fmt.Printf("Slide %d Shape %d: OUTSIDE pos (%d,%d) size (%d,%d) type=%d\n",
					i+1, si, sh.Left, sh.Top, sh.Width, sh.Height, sh.ShapeType)
			}
			if sh.Width < 0 || sh.Height < 0 {
				fmt.Printf("Slide %d Shape %d: NEGATIVE size (%d,%d) type=%d\n",
					i+1, si, sh.Width, sh.Height, sh.ShapeType)
			}
		}
	}
	fmt.Println("Position check complete.")
}
