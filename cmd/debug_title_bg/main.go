package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()

	// Check slides that use layout 4 (masterRef=2147483734)
	// These have white title text - need to understand what provides the dark bg
	for i, s := range slides {
		if s.GetMasterRef() != 2147483734 {
			continue
		}
		shapes := s.GetShapes()
		bg := s.GetBackground()

		// Find the title shape (usually first shape, large, white text, bold)
		for j, sh := range shapes {
			if len(sh.Paragraphs) == 0 {
				continue
			}
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.Color == "FFFFFF" && run.Bold && run.FontSize >= 2400 {
						fmt.Printf("Slide %d, Shape %d: WHITE BOLD title at (%d,%d) size=(%d,%d)\n",
							i+1, j, sh.Left, sh.Top, sh.Width, sh.Height)
						fmt.Printf("  Fill=%q NoFill=%v FillOpacity=%d FillColorRaw=0x%08X\n",
							sh.FillColor, sh.NoFill, sh.FillOpacity, sh.FillColorRaw)
						fmt.Printf("  Text=%q color=%s colorRaw=0x%08X\n",
							run.Text[:min(len(run.Text), 40)], run.Color, run.ColorRaw)
						fmt.Printf("  SlideBg: has=%v fill=%s imgIdx=%d\n",
							bg.HasBackground, bg.FillColor, bg.ImageIdx)

						// Check master for this slide
						if m, ok := masters[s.GetMasterRef()]; ok {
							fmt.Printf("  Master: bg=%v fill=%s imgIdx=%d\n",
								m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
							// Check if any master shape overlaps with this title
							for _, ms := range m.Shapes {
								if ms.IsImage && ms.ImageIdx >= 0 {
									// Check overlap
									imgRight := int64(ms.Left) + int64(ms.Width)
									imgBottom := int64(ms.Top) + int64(ms.Height)
									shapeCenterX := int64(sh.Left) + int64(sh.Width)/2
									shapeCenterY := int64(sh.Top) + int64(sh.Height)/2
									overlaps := shapeCenterX >= int64(ms.Left) && shapeCenterX <= imgRight &&
										shapeCenterY >= int64(ms.Top) && shapeCenterY <= imgBottom
									fmt.Printf("  Master img idx=%d at (%d,%d) size=(%d,%d) overlaps=%v\n",
										ms.ImageIdx, ms.Left, ms.Top, ms.Width, ms.Height, overlaps)
								}
							}
						}
						goto nextSlide
					}
				}
			}
		}
	nextSlide:
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
