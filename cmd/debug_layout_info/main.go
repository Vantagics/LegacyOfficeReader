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

	masters := p.GetMasters()
	slides := p.GetSlides()

	// Find unique master refs
	seen := make(map[uint32]bool)
	for _, s := range slides {
		ref := s.GetMasterRef()
		if seen[ref] {
			continue
		}
		seen[ref] = true

		fmt.Printf("MasterRef: %d\n", ref)
		if m, ok := masters[ref]; ok {
			fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
			fmt.Printf("  Background: has=%v fill=%s imgIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
			fmt.Printf("  Shapes: %d\n", len(m.Shapes))
			for i, sh := range m.Shapes {
				isConn := sh.ShapeType == 32 || sh.ShapeType == 33 || sh.ShapeType == 34
				fmt.Printf("    Shape %d: type=%d pos=(%d,%d) sz=(%d,%d) fill=%s noFill=%v isImage=%v imgIdx=%d isConn=%v\n",
					i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.NoFill, sh.IsImage, sh.ImageIdx, isConn)
			}
		}
	}

	// Check slide size
	w, h := p.GetSlideSize()
	fmt.Printf("\nSlide size: %d x %d\n", w, h)
}
