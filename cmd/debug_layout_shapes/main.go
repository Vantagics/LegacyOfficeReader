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

	// Build layout map
	masterRefToIdx := make(map[uint32]int)
	var layoutRefs []uint32
	for _, s := range slides {
		ref := s.GetMasterRef()
		if _, ok := masterRefToIdx[ref]; !ok {
			masterRefToIdx[ref] = len(layoutRefs)
			layoutRefs = append(layoutRefs, ref)
		}
	}

	// Check slide 13 (idx 12) layout
	sl := slides[12]
	ref := sl.GetMasterRef()
	layoutIdx := masterRefToIdx[ref]
	master := masters[ref]
	fmt.Printf("Slide 13 → layout %d (ref=%d)\n", layoutIdx, ref)
	fmt.Printf("  Master bg: fill=%q imgIdx=%d\n", master.Background.FillColor, master.Background.ImageIdx)
	fmt.Printf("  Master shapes: %d\n", len(master.Shapes))
	for si, sh := range master.Shapes {
		fmt.Printf("  MasterShape[%d] type=%d fill=%q noFill=%v img=%v imgIdx=%d pos=(%d,%d) size=(%d,%d)\n",
			si, sh.ShapeType, sh.FillColor, sh.NoFill, sh.IsImage, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height)
		if sh.IsText && len(sh.Paragraphs) > 0 {
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					t := run.Text
					if len(t) > 30 {
						t = t[:30] + "..."
					}
					fmt.Printf("    P[%d]R[%d] color=%q colorRaw=0x%08X sz=%d: %q\n",
						pi, ri, run.Color, run.ColorRaw, run.FontSize, t)
				}
			}
		}
	}

	// Also check slide 41 (idx 40) layout
	sl41 := slides[40]
	ref41 := sl41.GetMasterRef()
	layoutIdx41 := masterRefToIdx[ref41]
	master41 := masters[ref41]
	fmt.Printf("\nSlide 41 → layout %d (ref=%d)\n", layoutIdx41, ref41)
	fmt.Printf("  Master bg: fill=%q imgIdx=%d\n", master41.Background.FillColor, master41.Background.ImageIdx)
	fmt.Printf("  Master shapes: %d\n", len(master41.Shapes))
	for si, sh := range master41.Shapes {
		if sh.FillColor == "000000" || (sh.IsText && len(sh.Paragraphs) > 0) {
			fmt.Printf("  MasterShape[%d] type=%d fill=%q noFill=%v img=%v imgIdx=%d pos=(%d,%d) size=(%d,%d)\n",
				si, sh.ShapeType, sh.FillColor, sh.NoFill, sh.IsImage, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height)
			if sh.IsText {
				for pi, para := range sh.Paragraphs {
					for ri, run := range para.Runs {
						t := run.Text
						if len(t) > 30 {
							t = t[:30] + "..."
						}
						fmt.Printf("    P[%d]R[%d] color=%q colorRaw=0x%08X: %q\n",
							pi, ri, run.Color, run.ColorRaw, t)
					}
				}
			}
		}
	}
}
