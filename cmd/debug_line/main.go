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

	masters := p.GetMasters()
	m := masters[2147483734]
	
	fmt.Println("Master 2147483734 shapes:")
	for i, sh := range m.Shapes {
		fmt.Printf("  Shape %d: type=%d, lineColor=%s, lineWidth=%d, noLine=%v, lineDash=%d\n",
			i, sh.ShapeType, sh.LineColor, sh.LineWidth, sh.NoLine, sh.LineDash)
	}

	// Also check slide 4's connectors
	slides := p.GetSlides()
	if len(slides) >= 4 {
		s := slides[3] // slide 4 (0-indexed)
		fmt.Printf("\nSlide 4: %d shapes\n", len(s.GetShapes()))
		connCount := 0
		for _, sh := range s.GetShapes() {
			if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
				connCount++
				if connCount <= 5 {
					fmt.Printf("  Connector: type=%d, lineColor=%s, lineWidth=%d, noLine=%v, pos=(%d,%d) size=(%d,%d)\n",
						sh.ShapeType, sh.LineColor, sh.LineWidth, sh.NoLine, sh.Left, sh.Top, sh.Width, sh.Height)
				}
			}
		}
		fmt.Printf("  Total connectors: %d\n", connCount)
	}
}
