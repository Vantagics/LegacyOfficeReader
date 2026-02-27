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

	masters := p.GetMasters()
	slides := p.GetSlides()

	// Check line widths in masters
	fmt.Printf("=== Master Line Shapes ===\n")
	for ref, m := range masters {
		count := 0
		for _, s := range slides {
			if s.GetMasterRef() == ref {
				count++
			}
		}
		if count == 0 {
			continue
		}
		for i, sh := range m.Shapes {
			if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
				fmt.Printf("Master %d shape[%d]: type=%d lineColor=%s lineWidth=%d lineDash=%d\n",
					ref, i, sh.ShapeType, sh.LineColor, sh.LineWidth, sh.LineDash)
			}
		}
	}

	// Check line widths in slides
	fmt.Printf("\n=== Slide Line Width Distribution ===\n")
	lwDist := make(map[int32]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
				lwDist[sh.LineWidth]++
			}
		}
	}
	for lw, count := range lwDist {
		fmt.Printf("  lineWidth=%d: %d connectors\n", lw, count)
	}

	// Check line widths for non-connector shapes
	fmt.Printf("\n=== Non-Connector Line Width Distribution ===\n")
	lwDist2 := make(map[int32]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			if sh.ShapeType != 20 && !(sh.ShapeType >= 32 && sh.ShapeType <= 40) {
				if sh.LineWidth > 0 {
					lwDist2[sh.LineWidth]++
				}
			}
		}
	}
	for lw, count := range lwDist2 {
		fmt.Printf("  lineWidth=%d (%d EMU = %.1fpt): %d shapes\n", lw, lw, float64(lw)/12700.0, count)
	}

	// Check connector line widths in detail for slide 4
	fmt.Printf("\n=== Slide 4 Connectors ===\n")
	s4 := slides[3]
	for i, sh := range s4.GetShapes() {
		if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
			fmt.Printf("  [%d] type=%d lineColor=%s lineWidth=%d noLine=%v w=%d h=%d\n",
				i, sh.ShapeType, sh.LineColor, sh.LineWidth, sh.NoLine, sh.Width, sh.Height)
		}
	}
}
