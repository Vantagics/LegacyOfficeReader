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

	// Check the specific slides that had 0x04/0x05 flags
	targets := []int{31, 35, 36, 37, 38, 39, 45}
	for _, idx := range targets {
		if idx >= len(slides) {
			continue
		}
		s := slides[idx]
		ref := s.GetMasterRef()
		var scheme []string
		if m, ok := masters[ref]; ok {
			scheme = m.ColorScheme
		}
		shapes := s.GetShapes()
		fmt.Printf("\n=== Slide %d (scheme=%v) ===\n", idx+1, scheme)
		for si, sh := range shapes {
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					flag := run.ColorRaw >> 24
					if flag >= 0x01 && flag <= 0x07 {
						text := run.Text
						if len(text) > 30 {
							text = text[:30]
						}
						fmt.Printf("  Shape %d P%d R%d: raw=0x%08X resolved=%s flag=0x%02X text=%q\n",
							si, pi, ri, run.ColorRaw, run.Color, flag, text)
					}
				}
			}
		}
	}
}
