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
	masters := p.GetMasters()
	slideW, slideH := p.GetSlideSize()

	// For each slide, show what layout it maps to and what masterHasDarkBg would be
	masterRefToLayoutIdx := make(map[uint32]int)
	var layoutRefs []uint32
	for _, s := range slides {
		ref := s.GetMasterRef()
		if _, ok := masterRefToLayoutIdx[ref]; !ok {
			masterRefToLayoutIdx[ref] = len(layoutRefs)
			layoutRefs = append(layoutRefs, ref)
		}
	}

	fmt.Printf("Slide size: %d x %d\n", slideW, slideH)
	fmt.Printf("Layouts: %d\n\n", len(layoutRefs))

	for li, ref := range layoutRefs {
		m, ok := masters[ref]
		if !ok {
			fmt.Printf("Layout[%d] ref=%d: NO MASTER FOUND\n", li, ref)
			continue
		}

		// Check masterHasDarkBg logic
		hasBgImage := false
		hasDarkSolid := false
		hasFullPageImage := false
		for _, sh := range m.Shapes {
			if sh.IsImage && sh.ImageIdx >= 0 {
				if sh.Width > int32(float64(slideW)*0.7) && sh.Height > int32(float64(slideH)*0.7) {
					hasFullPageImage = true
				}
			}
		}
		if m.Background.ImageIdx >= 0 {
			hasBgImage = true
		}
		if m.Background.FillColor != "" && isDarkFillColor(m.Background.FillColor) {
			hasDarkSolid = true
		}

		masterHasDarkBg := hasBgImage || hasDarkSolid
		
		fmt.Printf("Layout[%d] ref=%d: bg.color=%s bg.imgIdx=%d hasBgImage=%v hasDarkSolid=%v masterHasDarkBg=%v hasFullPageImg=%v\n",
			li, ref, m.Background.FillColor, m.Background.ImageIdx, hasBgImage, hasDarkSolid, masterHasDarkBg, hasFullPageImage)
		fmt.Printf("  scheme: %v\n", m.ColorScheme)
		
		// Show which slides use this layout
		var slideNums []string
		for si, s := range slides {
			if s.GetMasterRef() == ref {
				slideNums = append(slideNums, fmt.Sprintf("%d", si+1))
			}
		}
		fmt.Printf("  slides: %s\n", strings.Join(slideNums, ", "))

		// Show shapes summary
		for si, sh := range m.Shapes {
			desc := fmt.Sprintf("type=%d %dx%d", sh.ShapeType, sh.Width, sh.Height)
			if sh.IsImage {
				desc += fmt.Sprintf(" IMG[%d]", sh.ImageIdx)
			}
			if sh.FillColor != "" {
				desc += " fill=" + sh.FillColor
			}
			fmt.Printf("  shape[%d]: %s\n", si, desc)
		}
		fmt.Println()
	}

	// Check specific slides for color issues
	fmt.Println("\n=== Color Issue Check ===")
	for si, slide := range slides {
		shapes := slide.GetShapes()
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
					// Check for potential visibility issues
					hasDarkFill := sh.FillColor != "" && isDarkFillColor(sh.FillColor)
					hasLightFill := sh.FillColor != "" && !isDarkFillColor(sh.FillColor)
					textIsDark := run.Color != "" && isDarkFillColor(run.Color)
					textIsLight := run.Color != "" && !isDarkFillColor(run.Color)
					textIsEmpty := run.Color == ""

					issue := ""
					if hasDarkFill && textIsDark {
						issue = "DARK_ON_DARK"
					} else if hasLightFill && textIsLight {
						issue = "LIGHT_ON_LIGHT"
					} else if hasDarkFill && textIsEmpty {
						issue = "NO_COLOR_ON_DARK"
					}

					if issue != "" {
						if len(t) > 40 {
							t = t[:40] + "..."
						}
						fmt.Printf("  Slide %d: %s fill=%s textColor=%s noFill=%v text=%q\n",
							si+1, issue, sh.FillColor, run.Color, sh.NoFill, t)
						break // one per shape
					}
				}
				break // one per shape
			}
		}
	}
}
