package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func isDarkFillColor(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	hexDigit := func(c byte) int {
		switch {
		case c >= '0' && c <= '9':
			return int(c - '0')
		case c >= 'A' && c <= 'F':
			return int(c-'A') + 10
		case c >= 'a' && c <= 'f':
			return int(c-'a') + 10
		}
		return 0
	}
	r := hexDigit(hex[0])*16 + hexDigit(hex[1])
	g := hexDigit(hex[2])*16 + hexDigit(hex[3])
	b := hexDigit(hex[4])*16 + hexDigit(hex[5])
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	return lum < 128
}

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	// Check specific problematic slides
	checkSlides := []int{3, 4, 8, 9, 17, 22, 25, 29, 33, 44, 45, 62, 63, 68, 69}
	for _, si := range checkSlides {
		if si >= len(slides) {
			continue
		}
		slide := slides[si]
		shapes := slide.GetShapes()
		scheme := slide.GetColorScheme()
		
		found := false
		for _, sh := range shapes {
			if !sh.IsText || len(sh.Paragraphs) == 0 {
				continue
			}
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t == "" {
						continue
					}
					
					hasDarkFill := sh.FillColor != "" && isDarkFillColor(sh.FillColor)
					hasLightFill := sh.FillColor != "" && !isDarkFillColor(sh.FillColor)
					textIsDark := run.Color != "" && isDarkFillColor(run.Color)
					textIsLight := run.Color != "" && !isDarkFillColor(run.Color)
					
					issue := ""
					if hasDarkFill && textIsDark {
						issue = "DARK_ON_DARK"
					} else if hasLightFill && textIsLight {
						issue = "LIGHT_ON_LIGHT"
					}
					
					if issue != "" && !found {
						found = true
						fmt.Printf("\n=== Slide %d (scheme=%v) ===\n", si+1, scheme)
					}
					if issue != "" {
						if len(t) > 30 {
							t = t[:30] + "..."
						}
						fmt.Printf("  %s fill=%s fillRaw=0x%08X text.color=%s text.colorRaw=0x%08X noFill=%v text=%q\n",
							issue, sh.FillColor, sh.FillColorRaw, run.Color, run.ColorRaw, sh.NoFill, t)
						break // one per shape
					}
				}
				if found {
					break
				}
			}
		}
	}
}
