package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func isDark(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	r := hexVal(hex[0])*16 + hexVal(hex[1])
	g := hexVal(hex[2])*16 + hexVal(hex[3])
	b := hexVal(hex[4])*16 + hexVal(hex[5])
	return (r*299+g*587+b*114)/1000 < 128
}

func hexVal(c byte) int {
	if c >= '0' && c <= '9' {
		return int(c - '0')
	}
	if c >= 'A' && c <= 'F' {
		return int(c-'A') + 10
	}
	if c >= 'a' && c <= 'f' {
		return int(c-'a') + 10
	}
	return 0
}

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
	targets := []int{31, 35, 36, 37, 38, 39, 45}
	for _, idx := range targets {
		if idx >= len(slides) {
			continue
		}
		s := slides[idx]
		shapes := s.GetShapes()
		fmt.Printf("\n=== Slide %d (all shapes) ===\n", idx+1)
		for si, sh := range shapes {
			text := ""
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := run.Text
					if len(t) > 20 {
						t = t[:20]
					}
					text += t
				}
			}
			fillInfo := fmt.Sprintf("fill=%q noFill=%v", sh.FillColor, sh.NoFill)
			if sh.IsImage {
				fillInfo = fmt.Sprintf("IMAGE idx=%d", sh.ImageIdx)
			}
			darkMark := ""
			if sh.FillColor != "" && isDark(sh.FillColor) {
				darkMark = " [DARK]"
			}
			fmt.Printf("  S%d: type=%d pos=(%d,%d) size=(%dx%d) %s%s text=%q\n",
				si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, fillInfo, darkMark, text)
		}
	}
}
