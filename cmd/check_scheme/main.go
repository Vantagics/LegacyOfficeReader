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

	// Check master 2147483734 (used by 56 slides)
	m, ok := masters[2147483734]
	if !ok {
		fmt.Println("Master 2147483734 not found")
		os.Exit(1)
	}

	fmt.Printf("Master 2147483734 color scheme: %v\n", m.ColorScheme)
	for i, sh := range m.Shapes {
		fmt.Printf("Shape[%d]: type=%d fill=%s fillRaw=0x%08X line=%s lineRaw=0x%08X\n",
			i, sh.ShapeType, sh.FillColor, sh.FillColorRaw, sh.LineColor, sh.LineColorRaw)
	}

	// Also check a slide's shapes
	slides := p.GetSlides()
	if len(slides) > 3 {
		s := slides[3] // slide 4
		fmt.Printf("\nSlide 4 (masterRef=%d):\n", s.GetMasterRef())
		for i, sh := range s.GetShapes() {
			if i > 5 {
				break
			}
			fmt.Printf("  Shape[%d]: type=%d fill=%s fillRaw=0x%08X line=%s lineRaw=0x%08X\n",
				i, sh.ShapeType, sh.FillColor, sh.FillColorRaw, sh.LineColor, sh.LineColorRaw)
		}
	}
}
