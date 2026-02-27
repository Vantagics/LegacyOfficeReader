package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	// Check slides where text color should be white (on dark/colored fills)
	// but might be incorrectly changed to dark
	whiteOnDark := 0
	darkOnLight := 0
	whiteOnLight := 0 // potential issues

	for si, slide := range slides {
		shapes := slide.GetShapes()
		for _, sh := range shapes {
			if len(sh.Paragraphs) == 0 {
				continue
			}
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.Text == "" {
						continue
					}
					isWhite := run.Color == "FFFFFF" || run.Color == ""
					hasDarkFill := sh.FillColor != "" && !sh.NoFill && isDark(sh.FillColor)
					hasLightFill := sh.FillColor != "" && !sh.NoFill && !isDark(sh.FillColor)
					isTransparent := sh.FillColor == "" || sh.NoFill

					if isWhite && hasDarkFill {
						whiteOnDark++
					} else if !isWhite && hasLightFill {
						darkOnLight++
					} else if isWhite && isTransparent {
						// White text on transparent shape - needs background
						_ = si // suppress unused
					} else if isWhite && hasLightFill && isNearWhite(sh.FillColor) {
						whiteOnLight++
						text := run.Text
						if len(text) > 30 {
							text = text[:30]
						}
						fmt.Printf("WARNING Slide %d: white text on light fill=%s: %q\n", si+1, sh.FillColor, text)
					}
				}
			}
		}
	}
	fmt.Printf("\nWhite on dark: %d, Dark on light: %d, White on light (issues): %d\n", whiteOnDark, darkOnLight, whiteOnLight)
}

func isDark(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	r := hexVal(hex[0])*16 + hexVal(hex[1])
	g := hexVal(hex[2])*16 + hexVal(hex[3])
	b := hexVal(hex[4])*16 + hexVal(hex[5])
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	return lum < 128
}

func isNearWhite(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	r := hexVal(hex[0])*16 + hexVal(hex[1])
	g := hexVal(hex[2])*16 + hexVal(hex[3])
	b := hexVal(hex[4])*16 + hexVal(hex[5])
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	return lum > 200
}

func hexVal(c byte) int {
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
