package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		panic(err)
	}
	defer f.Close()
	p, err := ppt.OpenReader(f)
	if err != nil {
		panic(err)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()

	// Check all masters for their shapes and backgrounds
	seen := make(map[uint32]bool)
	for _, s := range slides {
		ref := s.GetMasterRef()
		if seen[ref] {
			continue
		}
		seen[ref] = true

		m, ok := masters[ref]
		if !ok {
			continue
		}

		fmt.Printf("=== Master ref=%d ===\n", ref)
		fmt.Printf("  Background: has=%v fill=%q imgIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
		fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
		fmt.Printf("  Shapes: %d\n", len(m.Shapes))
		for si, sh := range m.Shapes {
			kind := "SHAPE"
			if sh.IsImage {
				kind = fmt.Sprintf("IMAGE(idx=%d)", sh.ImageIdx)
			} else if sh.IsText {
				kind = "TEXT"
			}
			fmt.Printf("  [%d] type=%d %s pos=(%d,%d) sz=(%d,%d) fill=%q noFill=%v\n",
				si, sh.ShapeType, kind, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.NoFill)
		}

		// Check which slides use this master
		slideNums := []int{}
		for i, s2 := range slides {
			if s2.GetMasterRef() == ref {
				slideNums = append(slideNums, i+1)
				if len(slideNums) > 10 {
					break
				}
			}
		}
		fmt.Printf("  Used by slides: %v\n\n", slideNums)
	}

	// Specifically check slide 4's context
	fmt.Println("=== Slide 4 detailed ===")
	s4 := slides[3]
	fmt.Printf("MasterRef: %d\n", s4.GetMasterRef())
	bg := s4.GetBackground()
	fmt.Printf("Slide bg: has=%v fill=%q imgIdx=%d\n", bg.HasBackground, bg.FillColor, bg.ImageIdx)

	// Check if the title area has any shape behind it
	shapes := s4.GetShapes()
	titleTop := int32(396875)
	titleBottom := titleTop + int32(590550)
	fmt.Printf("Title area: y=%d to y=%d\n", titleTop, titleBottom)
	for si, sh := range shapes {
		if sh.IsImage && sh.Top <= titleTop && sh.Top+sh.Height >= titleBottom {
			fmt.Printf("  Image shape[%d] covers title: pos=(%d,%d) sz=(%d,%d) imgIdx=%d\n",
				si, sh.Left, sh.Top, sh.Width, sh.Height, sh.ImageIdx)
		}
		if !sh.IsImage && !sh.IsText && sh.FillColor != "" && sh.Top <= titleTop && sh.Top+sh.Height >= titleBottom {
			fmt.Printf("  Filled shape[%d] covers title: pos=(%d,%d) sz=(%d,%d) fill=%q\n",
				si, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor)
		}
	}
}
