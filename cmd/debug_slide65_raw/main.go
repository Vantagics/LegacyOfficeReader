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
	slide := slides[64]
	shapes := slide.GetShapes()

	// Focus on Shape[1] (main text), Shape[2] (yellow box), Shape[3] (tips), Shape[4] (bottom text)
	for i := 0; i < len(shapes); i++ {
		s := shapes[i]
		hasText := false
		for _, p := range s.Paragraphs {
			for _, r := range p.Runs {
				if len(r.Text) > 0 {
					hasText = true
					break
				}
			}
			if hasText {
				break
			}
		}
		if !hasText {
			continue
		}

		fmt.Printf("=== Shape[%d] type=%d pos=(%d,%d) size=(%d,%d) ===\n", i, s.ShapeType, s.Left, s.Top, s.Width, s.Height)
		fmt.Printf("  FillColor=%q NoFill=%v\n", s.FillColor, s.NoFill)
		fmt.Printf("  TextMargins: L=%d T=%d R=%d B=%d\n", s.TextMarginLeft, s.TextMarginTop, s.TextMarginRight, s.TextMarginBottom)
		for pi, para := range s.Paragraphs {
			fmt.Printf("  Para[%d]: lineSpacing=%d\n", pi, para.LineSpacing)
			for ri, run := range para.Runs {
				text := run.Text
				if len(text) > 50 {
					text = text[:50] + "..."
				}
				fmt.Printf("    Run[%d]: sz=%d color=%q colorRaw=0x%08X bold=%v font=%q text=%q\n",
					ri, run.FontSize, run.Color, run.ColorRaw, run.Bold, run.FontName, text)
			}
		}
		fmt.Println()
	}

	// Also check master color scheme
	masters := p.GetMasters()
	fmt.Println("=== Master Color Schemes ===")
	for ref, m := range masters {
		fmt.Printf("Master %d: scheme=%v\n", ref, m.ColorScheme)
		break // just show first
	}
}
