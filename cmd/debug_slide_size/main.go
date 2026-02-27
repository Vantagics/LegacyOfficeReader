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
	w, h := p.GetSlideSize()
	fmt.Printf("Slide size: %d x %d EMU (%.1f x %.1f inches)\n", w, h, float64(w)/914400, float64(h)/914400)

	// Check isFullPageImage for layout 4 shapes
	masters := p.GetMasters()
	slides := p.GetSlides()
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

	// Layout 4 = index 3 (ref=2147483734)
	if len(layoutRefs) > 3 {
		ref := layoutRefs[3]
		m := masters[ref]
		fmt.Printf("\nLayout 4 (ref=%d) shapes:\n", ref)
		for i, sh := range m.Shapes {
			isFullPage := sh.Width > int32(float64(w)*0.7) && sh.Height > int32(float64(h)*0.7)
			fmt.Printf("  [%d] type=%d size=(%d,%d) img=%v imgIdx=%d isFullPage=%v (w>%.0f=%v h>%.0f=%v)\n",
				i, sh.ShapeType, sh.Width, sh.Height, sh.IsImage, sh.ImageIdx,
				isFullPage,
				float64(w)*0.7, sh.Width > int32(float64(w)*0.7),
				float64(h)*0.7, sh.Height > int32(float64(h)*0.7))
		}
	}

	// Check all layouts
	for i, ref := range layoutRefs {
		m := masters[ref]
		fmt.Printf("\nLayout %d (ref=%d):\n", i+1, ref)
		for j, sh := range m.Shapes {
			if sh.IsImage {
				isFullPage := sh.Width > int32(float64(w)*0.7) && sh.Height > int32(float64(h)*0.7)
				fmt.Printf("  [%d] img imgIdx=%d size=(%d,%d) isFullPage=%v\n",
					j, sh.ImageIdx, sh.Width, sh.Height, isFullPage)
			}
		}
	}
}
