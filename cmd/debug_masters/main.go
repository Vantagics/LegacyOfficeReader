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
	images := p.GetImages()
	w, h := p.GetSlideSize()

	fmt.Printf("Slides: %d, Images: %d, Masters: %d\n", len(slides), len(images), len(masters))
	fmt.Printf("Slide size: %d x %d EMU\n", w, h)

	// Print master info
	fmt.Println("\n=== Masters ===")
	for ref, m := range masters {
		fmt.Printf("Master ref=%d: bg=%v, shapes=%d, colorScheme=%v\n",
			ref, m.Background.HasBackground, len(m.Shapes), m.ColorScheme)
		for i, sh := range m.Shapes {
			text := ""
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					text += r.Text
				}
			}
			if text != "" {
				fmt.Printf("  Shape %d: type=%d, isText=%v, isImage=%v, text=%q\n", i, sh.ShapeType, sh.IsText, sh.IsImage, text)
			} else {
				fmt.Printf("  Shape %d: type=%d, isText=%v, isImage=%v, imgIdx=%d, pos=(%d,%d) size=(%d,%d)\n",
					i, sh.ShapeType, sh.IsText, sh.IsImage, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height)
			}
		}
	}

	// Print slide masterRef distribution
	fmt.Println("\n=== Slide masterRef distribution ===")
	refCount := make(map[uint32]int)
	for _, s := range slides {
		refCount[s.GetMasterRef()]++
	}
	for ref, count := range refCount {
		fmt.Printf("  masterRef=%d: %d slides\n", ref, count)
	}

	// Print first few slides' details
	fmt.Println("\n=== First 5 slides ===")
	for i := 0; i < 5 && i < len(slides); i++ {
		s := slides[i]
		fmt.Printf("Slide %d: masterRef=%d, shapes=%d, bg=%v, layout=%d\n",
			i+1, s.GetMasterRef(), len(s.GetShapes()), s.GetBackground().HasBackground, s.GetLayoutType())
		for j, sh := range s.GetShapes() {
			text := ""
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					text += r.Text
				}
			}
			if len(text) > 80 {
				text = text[:80] + "..."
			}
			if text != "" {
				fmt.Printf("  Shape %d: type=%d, text=%q\n", j, sh.ShapeType, text)
			} else if sh.IsImage {
				fmt.Printf("  Shape %d: image idx=%d, pos=(%d,%d) size=(%d,%d)\n", j, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height)
			}
		}
	}
}
