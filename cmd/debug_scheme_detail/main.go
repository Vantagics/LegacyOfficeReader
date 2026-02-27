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
	// Check slides 32, 36-40, 46
	targets := []int{31, 35, 36, 37, 38, 39, 45} // 0-indexed
	for _, idx := range targets {
		if idx >= len(slides) {
			continue
		}
		s := slides[idx]
		shapes := s.GetShapes()
		fmt.Printf("\n=== Slide %d ===\n", idx+1)
		for si, sh := range shapes {
			hasSpecialFlag := false
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					flag := run.ColorRaw & 0xFF000000
					if flag == 0x04000000 || flag == 0x05000000 {
						hasSpecialFlag = true
					}
				}
			}
			if !hasSpecialFlag {
				continue
			}
			fmt.Printf("  Shape %d: type=%d pos=(%d,%d) size=(%dx%d) fill=%q noFill=%v fillRaw=0x%08X\n",
				si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
				sh.FillColor, sh.NoFill, sh.FillColorRaw)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					text := run.Text
					if len(text) > 40 {
						text = text[:40]
					}
					fmt.Printf("    P%d R%d: color=%s raw=0x%08X bold=%v size=%d text=%q\n",
						pi, ri, run.Color, run.ColorRaw, run.Bold, run.FontSize, text)
				}
			}
		}
	}
}
