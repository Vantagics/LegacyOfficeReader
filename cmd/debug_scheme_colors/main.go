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

	// Check scheme[0] colors across slides to see if embedded RGB differs from scheme
	checkSlides := []int{0, 3, 4, 7, 12, 20, 40, 57}
	for _, idx := range checkSlides {
		if idx >= len(slides) {
			continue
		}
		s := slides[idx]
		shapes := s.GetShapes()
		scheme := s.GetColorScheme()
		fmt.Printf("=== Slide %d (scheme[0]=%s, scheme[1]=%s) ===\n", idx+1, scheme[0], scheme[1])

		for si, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.ColorRaw != 0 {
						flag := run.ColorRaw & 0xFF000000
						if flag == 0xFE000000 || flag == 0x08000000 {
							schemeIdx := run.ColorRaw & 0xFF
							embeddedR := uint8(run.ColorRaw & 0xFF)
							embeddedG := uint8((run.ColorRaw >> 8) & 0xFF)
							embeddedB := uint8((run.ColorRaw >> 16) & 0xFF)
							embeddedRGB := fmt.Sprintf("%02X%02X%02X", embeddedR, embeddedG, embeddedB)
							resolved := run.Color
							text := run.Text
							if len([]rune(text)) > 20 {
								text = string([]rune(text)[:20]) + "..."
							}
							if schemeIdx == 0 || schemeIdx == 1 {
								fmt.Printf("  Shape[%d] raw=0x%08X idx=%d embedded=%s resolved=%s text=%q\n",
									si, run.ColorRaw, schemeIdx, embeddedRGB, resolved, text)
							}
						}
					}
				}
			}
		}
	}
}
