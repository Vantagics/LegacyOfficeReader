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
	w, h := p.GetSlideSize()
	fmt.Printf("Slide size: %d x %d\n", w, h)

	// Show details for the masters used by slides
	refs := []uint32{2147483734, 2147483728, 2147483735, 2147483737, 2147483745, 2147483700, 2147483730}
	for _, ref := range refs {
		m, ok := masters[ref]
		if !ok {
			continue
		}
		fmt.Printf("\n=== Master %d ===\n", ref)
		fmt.Printf("Background: fill=%s, imgIdx=%d\n", m.Background.FillColor, m.Background.ImageIdx)
		for i, sh := range m.Shapes {
			text := ""
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					text += r.Text
				}
			}
			fmt.Printf("  Shape %d: type=%d, isText=%v, isImage=%v, imgIdx=%d\n", i, sh.ShapeType, sh.IsText, sh.IsImage, sh.ImageIdx)
			fmt.Printf("    pos=(%d,%d) size=(%d,%d)\n", sh.Left, sh.Top, sh.Width, sh.Height)
			fmt.Printf("    fill=%s, noFill=%v, line=%s, noLine=%v\n", sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine)
			if text != "" {
				if len(text) > 80 {
					text = text[:80] + "..."
				}
				fmt.Printf("    text=%q\n", text)
			}
		}
	}
}
