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
	s := slides[7] // slide 8
	shapes := s.GetShapes()
	for i, sh := range shapes {
		if len(sh.GeoVertices) == 0 {
			continue
		}
		fmt.Printf("shape[%d]: type=%d pos=(%d,%d) size=(%d,%d)\n", i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height)
		fmt.Printf("  Geo: left=%d top=%d right=%d bottom=%d\n", sh.GeoLeft, sh.GeoTop, sh.GeoRight, sh.GeoBottom)
		fmt.Printf("  Vertices: %d, Segments: %d\n", len(sh.GeoVertices), len(sh.GeoSegments))
		fmt.Printf("  Fill: %q NoFill=%v\n", sh.FillColor, sh.NoFill)

		// Show first few vertices
		for j := 0; j < 5 && j < len(sh.GeoVertices); j++ {
			v := sh.GeoVertices[j]
			fmt.Printf("    v[%d]: (%d, %d)\n", j, v.X, v.Y)
		}
		// Show segments
		for j, seg := range sh.GeoSegments {
			fmt.Printf("    seg[%d]: type=%d count=%d\n", j, seg.SegType, seg.Count)
		}
		fmt.Println()
	}
}
