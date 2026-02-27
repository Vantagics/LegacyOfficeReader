package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Parse error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()
	// Check slides 13, 15, 41, 71 (0-indexed: 12, 14, 40, 70)
	checkSlides := []int{12, 14, 40, 70}
	for _, idx := range checkSlides {
		if idx >= len(slides) {
			continue
		}
		sl := slides[idx]
		ref := sl.GetMasterRef()
		master := masters[ref]
		fmt.Printf("=== SLIDE %d (masterRef=%d) ===\n", idx+1, ref)
		fmt.Printf("  ColorScheme: %v\n", master.ColorScheme)
		fmt.Printf("  SlideBg: fill=%q imgIdx=%d\n", sl.GetBackground().FillColor, sl.GetBackground().ImageIdx)
		for si, sh := range sl.GetShapes() {
			if !sh.IsText || len(sh.Paragraphs) == 0 {
				continue
			}
			hasText := false
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.Text != "" {
						hasText = true
						break
					}
				}
			}
			if !hasText {
				continue
			}
			fmt.Printf("  Shape[%d] type=%d fill=%q noFill=%v fillRaw=0x%08X opacity=%d\n",
				si, sh.ShapeType, sh.FillColor, sh.NoFill, sh.FillColorRaw, sh.FillOpacity)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					if run.Text == "" {
						continue
					}
					t := run.Text
					if len(t) > 40 {
						t = t[:40] + "..."
					}
					fmt.Printf("    P[%d]R[%d] color=%q colorRaw=0x%08X sz=%d: %q\n",
						pi, ri, run.Color, run.ColorRaw, run.FontSize, t)
				}
			}
		}
	}
}
