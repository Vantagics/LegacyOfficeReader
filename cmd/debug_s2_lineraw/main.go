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
	
	// Master color scheme for this slide
	masters := p.GetMasters()
	m := masters[s.GetMasterRef()]
	fmt.Printf("Master ref=%d, colorScheme=%v\n\n", s.GetMasterRef(), m.ColorScheme)

	for i, sh := range shapes {
		if sh.LineColor == "" {
			continue
		}
		text := ""
		for _, p := range sh.Paragraphs {
			for _, r := range p.Runs {
				text += r.Text
			}
		}
		fmt.Printf("shape[%d]: type=%d line=%q lineRaw=0x%08X noLine=%v lineW=%d text=%q\n",
			i, sh.ShapeType, sh.LineColor, sh.LineColorRaw, sh.NoLine, sh.LineWidth, text)
	}
}
