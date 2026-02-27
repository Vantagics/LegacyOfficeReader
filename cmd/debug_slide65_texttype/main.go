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
	if len(slides) < 65 {
		fmt.Println("Not enough slides")
		return
	}

	slide := slides[64] // 0-indexed, slide 65
	fmt.Printf("=== Slide 65 ===\n")
	fmt.Printf("MasterRef: %d\n", slide.GetMasterRef())
	fmt.Printf("ColorScheme: %v\n", slide.GetColorScheme())
	fmt.Printf("DefaultTextStyles:\n")
	for i, s := range slide.GetDefaultTextStyles() {
		fmt.Printf("  Level %d: sz=%d font=%q color=%s colorRaw=0x%08X\n", i, s.FontSize, s.FontName, s.Color, s.ColorRaw)
	}

	// Check text type styles from master
	masters := p.GetMasters()
	if m, ok := masters[slide.GetMasterRef()]; ok {
		fmt.Printf("\nMaster TextTypeStyles (including env):\n")
		for tt, styles := range m.TextTypeStyles {
			fmt.Printf("  TextType %d:\n", tt)
			for i, s := range styles {
				if s.FontSize > 0 || s.Color != "" {
					fmt.Printf("    Level %d: sz=%d font=%q color=%s colorRaw=0x%08X bold=%v\n", i, s.FontSize, s.FontName, s.Color, s.ColorRaw, s.Bold)
				}
			}
		}
	}

	fmt.Printf("\nShapes:\n")
	for i, sh := range slide.GetShapes() {
		fmt.Printf("  Shape[%d]: type=%d textType=%d fill=%q noFill=%v pos=(%d,%d) size=(%d,%d)\n",
			i, sh.ShapeType, sh.TextType, sh.FillColor, sh.NoFill, sh.Left, sh.Top, sh.Width, sh.Height)
		for pi, para := range sh.Paragraphs {
			for ri, run := range para.Runs {
				text := run.Text
				if len(text) > 50 {
					text = text[:50] + "..."
				}
				fmt.Printf("    P%d R%d: sz=%d color=%s colorRaw=0x%08X font=%q text=%q\n",
					pi, ri, run.FontSize, run.Color, run.ColorRaw, run.FontName, text)
			}
		}
	}
}
