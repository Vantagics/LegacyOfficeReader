package main

import (
	"fmt"
	"os"
	"strings"

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

	// Build unique layout list from master refs (same as pptconv)
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

	fmt.Printf("Total masters: %d\n", len(masters))
	fmt.Printf("Unique layouts: %d\n", len(layoutRefs))

	for i, ref := range layoutRefs {
		m, ok := masters[ref]
		if !ok {
			fmt.Printf("\nLayout %d (ref=%d): NOT FOUND in masters\n", i+1, ref)
			continue
		}

		// Count slides using this layout
		slideCount := 0
		for _, s := range slides {
			if s.GetMasterRef() == ref {
				slideCount++
			}
		}

		fmt.Printf("\nLayout %d (ref=%d, %d slides):\n", i+1, ref, slideCount)
		fmt.Printf("  Background: has=%v fill=%q imgIdx=%d\n",
			m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
		fmt.Printf("  Color scheme: %v\n", m.ColorScheme)
		fmt.Printf("  Shapes: %d\n", len(m.Shapes))

		for j, sh := range m.Shapes {
			text := ""
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t != "" {
						text += t + " "
					}
				}
			}
			if len([]rune(text)) > 50 {
				text = string([]rune(text)[:50]) + "..."
			}

			fmt.Printf("  [%d] type=%d pos=(%d,%d) size=(%d,%d) img=%v imgIdx=%d fill=%q noFill=%v line=%q text=%q\n",
				j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
				sh.IsImage, sh.ImageIdx, sh.FillColor, sh.NoFill, sh.LineColor, text)
		}
	}
}
