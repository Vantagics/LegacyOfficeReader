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
	masters := p.GetMasters()

	// Build layout mapping (same as mapFormattedSlides)
	masterRefToLayoutIdx := map[uint32]int{}
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

	fmt.Printf("Layout mapping (%d layouts):\n", len(layoutRefs))
	for i, ref := range layoutRefs {
		m, found := masters[ref]
		slideCount := 0
		for _, s := range slides {
			if s.GetMasterRef() == ref {
				slideCount++
			}
		}
		fmt.Printf("  Layout %d: masterRef=%d found=%v slides=%d shapes=%d bg=(has=%v,color=%s,img=%d)\n",
			i+1, ref, found, slideCount, len(m.Shapes), m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
		for j, sh := range m.Shapes {
			fmt.Printf("    Shape[%d]: type=%d isText=%v isImage=%v imgIdx=%d pos=(%d,%d) size=(%d,%d)\n",
				j, sh.ShapeType, sh.IsText, sh.IsImage, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height)
		}
	}
}
