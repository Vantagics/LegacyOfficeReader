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
	slide := slides[7]
	shapes := slide.GetShapes()

	for i, sh := range shapes {
		if sh.ShapeType != 0 {
			continue
		}
		fmt.Printf("Shape %d: type=0 fill=%q noFill=%v fillBoolsRaw=? pos=(%d,%d) sz=(%d,%d) verts=%d segs=%d\n",
			i, sh.FillColor, sh.NoFill, sh.Left, sh.Top, sh.Width, sh.Height,
			len(sh.GeoVertices), len(sh.GeoSegments))
	}
}
