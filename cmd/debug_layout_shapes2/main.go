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

	// Check what master each slide uses
	fmt.Println("=== Slide to Master mapping ===")
	masterRefToIdx := make(map[uint32]int)
	layoutIdx := 0
	for _, s := range slides {
		ref := s.GetMasterRef()
		if _, ok := masterRefToIdx[ref]; !ok {
			masterRefToIdx[ref] = layoutIdx
			layoutIdx++
		}
	}

	// Print layout details for each unique master
	for ref, idx := range masterRefToIdx {
		m, ok := masters[ref]
		if !ok {
			fmt.Printf("Layout %d (ref=%d): NOT FOUND\n", idx, ref)
			continue
		}
		fmt.Printf("\n=== Layout %d (ref=%d) ===\n", idx, ref)
		fmt.Printf("  Background: has=%v fill=%q imgIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
		fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
		fmt.Printf("  Shapes: %d\n", len(m.Shapes))
		for si, sh := range m.Shapes {
			desc := ""
			if sh.IsImage {
				desc = fmt.Sprintf("IMAGE imgIdx=%d", sh.ImageIdx)
			} else if sh.IsText {
				text := ""
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						text += run.Text
					}
				}
				desc = fmt.Sprintf("TEXT %q", truncate(text, 30))
			} else {
				desc = "SHAPE"
			}
			fmt.Printf("  [%d] type=%d pos=(%d,%d) sz=(%d,%d) fill=%q line=%q noFill=%v noLine=%v %s\n",
				si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
				sh.FillColor, sh.LineColor, sh.NoFill, sh.NoLine, desc)
		}
	}

	// Check specific slides for their title shape context
	fmt.Println("\n=== Slide 4 title shape analysis ===")
	if len(slides) >= 4 {
		s := slides[3]
		shapes := s.GetShapes()
		ref := s.GetMasterRef()
		fmt.Printf("Slide 4 masterRef=%d layoutIdx=%d\n", ref, masterRefToIdx[ref])
		if len(shapes) > 0 {
			sh := shapes[0]
			fmt.Printf("Shape[0]: type=%d pos=(%d,%d) sz=(%d,%d)\n", sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height)
			fmt.Printf("  fill=%q noFill=%v fillOpacity=%d fillColorRaw=0x%08X\n", sh.FillColor, sh.NoFill, sh.FillOpacity, sh.FillColorRaw)
			fmt.Printf("  line=%q noLine=%v\n", sh.LineColor, sh.NoLine)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					fmt.Printf("  Para[%d] Run[%d]: color=%q colorRaw=0x%08X fontSize=%d text=%q\n",
						pi, ri, run.Color, run.ColorRaw, run.FontSize, truncate(run.Text, 40))
				}
			}
		}
	}

	// Check slide 1 title
	fmt.Println("\n=== Slide 1 title shape analysis ===")
	if len(slides) >= 1 {
		s := slides[0]
		shapes := s.GetShapes()
		ref := s.GetMasterRef()
		fmt.Printf("Slide 1 masterRef=%d layoutIdx=%d\n", ref, masterRefToIdx[ref])
		for si, sh := range shapes {
			fmt.Printf("Shape[%d]: type=%d pos=(%d,%d) sz=(%d,%d)\n", si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height)
			fmt.Printf("  fill=%q noFill=%v fillOpacity=%d fillColorRaw=0x%08X\n", sh.FillColor, sh.NoFill, sh.FillOpacity, sh.FillColorRaw)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					fmt.Printf("  Para[%d] Run[%d]: color=%q colorRaw=0x%08X fontSize=%d text=%q\n",
						pi, ri, run.Color, run.ColorRaw, run.FontSize, truncate(run.Text, 40))
				}
			}
		}
	}

	// Check slide 5 subtitle shape (fontSize=0)
	fmt.Println("\n=== Slide 5 subtitle shape (fontSize=0) ===")
	if len(slides) >= 5 {
		s := slides[4]
		shapes := s.GetShapes()
		for si, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.FontSize == 0 {
						fmt.Printf("Shape[%d]: type=%d pos=(%d,%d) sz=(%d,%d) fontSize=0 text=%q\n",
							si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, truncate(run.Text, 40))
						break
					}
				}
			}
		}
	}
}

func truncate(s string, n int) string {
	r := []rune(s)
	if len(r) > n {
		return string(r[:n]) + "..."
	}
	return s
}
