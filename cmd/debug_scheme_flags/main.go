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

	for i, s := range slides {
		ref := s.GetMasterRef()
		var scheme []string
		if m, ok := masters[ref]; ok {
			scheme = m.ColorScheme
		}
		shapes := s.GetShapes()
		for si, sh := range shapes {
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					if run.ColorRaw == 0 {
						continue
					}
					flag := run.ColorRaw & 0xFF000000
					// Show all non-standard flags (not 0x08, 0xFE, 0x00)
					if flag != 0x08000000 && flag != 0xFE000000 && flag != 0x00000000 {
						resolved := ""
						idx := int(run.ColorRaw & 0xFF)
						if idx < len(scheme) {
							resolved = scheme[idx]
						}
						text := run.Text
						if len(text) > 30 {
							text = text[:30]
						}
						fmt.Printf("Slide %d, Shape %d, Para %d, Run %d: flag=0x%08X color=%s idx=%d scheme[idx]=%s text=%q\n",
							i+1, si, pi, ri, run.ColorRaw, run.Color, idx, resolved, text)
					}
				}
			}
			// Also check fill colors
			if sh.FillColorRaw != 0 {
				flag := sh.FillColorRaw & 0xFF000000
				if flag != 0x08000000 && flag != 0xFE000000 && flag != 0x00000000 {
					idx := int(sh.FillColorRaw & 0xFF)
					resolved := ""
					if idx < len(scheme) {
						resolved = scheme[idx]
					}
					fmt.Printf("Slide %d, Shape %d FILL: flag=0x%08X color=%s idx=%d scheme[idx]=%s\n",
						i+1, si, sh.FillColorRaw, sh.FillColor, idx, resolved)
				}
			}
		}
	}

	// Also print all unique schemes
	fmt.Println("\n--- Color Schemes ---")
	seen := map[uint32]bool{}
	for _, s := range slides {
		ref := s.GetMasterRef()
		if seen[ref] {
			continue
		}
		seen[ref] = true
		if m, ok := masters[ref]; ok {
			fmt.Printf("Master ref=%d scheme=%v\n", ref, m.ColorScheme)
		}
	}
}
