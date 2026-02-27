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
	// MasterRef 2147483734 is the main layout
	ref := uint32(2147483734)
	m, ok := masters[ref]
	if !ok {
		fmt.Println("Master not found")
		return
	}

	fmt.Printf("Master %d: %d shapes, colorScheme=%v\n", ref, len(m.Shapes), m.ColorScheme)
	fmt.Printf("Background: has=%v fill=%s imageIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)

	for i, sh := range m.Shapes {
		text := ""
		for _, p := range sh.Paragraphs {
			for _, r := range p.Runs {
				if r.Text != "" {
					text = r.Text
					if len(text) > 30 {
						text = text[:30] + "..."
					}
					break
				}
			}
			if text != "" {
				break
			}
		}
		fmt.Printf("Shape %d: type=%d isImage=%v imgIdx=%d fill=%s noFill=%v pos=(%d,%d) sz=(%d,%d) text=%q\n",
			i, sh.ShapeType, sh.IsImage, sh.ImageIdx, sh.FillColor, sh.NoFill,
			sh.Left, sh.Top, sh.Width, sh.Height, text)
		if sh.ShapeType == 0 && len(sh.GeoVertices) > 0 {
			fmt.Printf("  -> freeform: %d vertices, %d segments\n", len(sh.GeoVertices), len(sh.GeoSegments))
		}
	}
}
