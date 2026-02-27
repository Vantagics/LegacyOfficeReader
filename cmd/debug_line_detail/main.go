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
	slides := p.GetSlides()

	// Check layout 4 connector line details
	masterRefToLayoutIdx := make(map[uint32]int)
	var layoutRefs []uint32
	for _, s := range slides {
		ref := s.GetMasterRef()
		if _, ok := masterRefToLayoutIdx[ref]; ok {
			continue
		}
		idx := len(layoutRefs)
		masterRefToLayoutIdx[ref] = idx
		layoutRefs = append(layoutRefs, ref)
	}

	for i, ref := range layoutRefs {
		m := masters[ref]
		for j, sh := range m.Shapes {
			if sh.ShapeType == 20 || sh.LineColor != "" {
				fmt.Printf("Layout %d Shape %d: type=%d lineColor=%q lineWidth=%d lineDash=%d noLine=%v\n",
					i+1, j, sh.ShapeType, sh.LineColor, sh.LineWidth, sh.LineDash, sh.NoLine)
			}
		}
	}

	// Also check a few slide shapes for line details
	fmt.Println("\n--- Slide 4 connector lines ---")
	shapes := slides[3].GetShapes()
	for j, sh := range shapes {
		if sh.ShapeType == 20 {
			fmt.Printf("  Shape %d: lineColor=%q lineWidth=%d lineDash=%d noLine=%v pos=(%d,%d) size=(%d,%d)\n",
				j, sh.LineColor, sh.LineWidth, sh.LineDash, sh.NoLine, sh.Left, sh.Top, sh.Width, sh.Height)
			if j > 35 {
				break
			}
		}
	}
}
