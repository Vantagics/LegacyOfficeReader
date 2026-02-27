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
	s := slides[1] // slide 2
	shapes := s.GetShapes()

	for i, sh := range shapes {
		text := ""
		for _, p := range sh.Paragraphs {
			for _, r := range p.Runs {
				text += r.Text
			}
		}
		if len(text) > 40 {
			text = text[:40]
		}

		fmt.Printf("shape[%d]: type=%d verts=%d segs=%d geo=(%d,%d,%d,%d) text=%q\n",
			i, sh.ShapeType, len(sh.GeoVertices), len(sh.GeoSegments),
			sh.GeoLeft, sh.GeoTop, sh.GeoRight, sh.GeoBottom, text)
		if len(sh.GeoVertices) > 0 {
			for j := 0; j < len(sh.GeoVertices) && j < 5; j++ {
				v := sh.GeoVertices[j]
				fmt.Printf("  v[%d]: (%d, %d)\n", j, v.X, v.Y)
			}
			for j, seg := range sh.GeoSegments {
				fmt.Printf("  seg[%d]: type=%d count=%d\n", j, seg.SegType, seg.Count)
			}
		}
	}
}
