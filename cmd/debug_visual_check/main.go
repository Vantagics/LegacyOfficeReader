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

	// Issue 1: Check for runs with sz=0 that need estimation
	fmt.Println("=== Runs with FontSize=0 (need estimation) ===")
	for i, s := range slides {
		for j, sh := range s.GetShapes() {
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					if run.FontSize == 0 && run.Text != "" {
						text := run.Text
						if len(text) > 30 {
							text = text[:30]
						}
						fmt.Printf("  Slide %d shape %d p%d.r%d: sz=0 shapeH=%d %q\n",
							i+1, j, pi, ri, sh.Height, text)
					}
				}
			}
		}
	}

	// Issue 2: Check for runs with empty color
	fmt.Println("\n=== Runs with empty Color (need scheme resolution) ===")
	for i, s := range slides {
		cs := s.GetColorScheme()
		for j, sh := range s.GetShapes() {
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					if run.Color == "" && run.Text != "" {
						text := run.Text
						if len(text) > 30 {
							text = text[:30]
						}
						fmt.Printf("  Slide %d shape %d p%d.r%d: noColor fill=%q scheme[1]=%s %q\n",
							i+1, j, pi, ri, sh.FillColor, safeIdx(cs, 1), text)
					}
				}
			}
		}
	}

	// Issue 3: Check master shapes for scheme color issues
	fmt.Println("\n=== Master Shape Colors ===")
	for ref, m := range masters {
		for i, sh := range m.Shapes {
			if sh.FillColor != "" {
				fmt.Printf("  Master %d shape %d: fill=%s raw=0x%08X\n", ref, i, sh.FillColor, sh.FillColorRaw)
			}
			if sh.LineColor != "" {
				fmt.Printf("  Master %d shape %d: line=%s raw=0x%08X\n", ref, i, sh.LineColor, sh.LineColorRaw)
			}
		}
	}

	// Issue 4: Check for shapes with negative coordinates
	fmt.Println("\n=== Shapes with negative coordinates ===")
	negCount := 0
	for i, s := range slides {
		for j, sh := range s.GetShapes() {
			if sh.Left < 0 || sh.Top < 0 {
				negCount++
				if negCount <= 10 {
					fmt.Printf("  Slide %d shape %d: pos=(%d,%d) size=(%d,%d)\n",
						i+1, j, sh.Left, sh.Top, sh.Width, sh.Height)
				}
			}
		}
	}
	fmt.Printf("  Total shapes with negative coords: %d\n", negCount)

	// Issue 5: Check for shapes with very small font sizes (< 600)
	fmt.Println("\n=== Shapes with very small font sizes ===")
	for i, s := range slides {
		for j, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.FontSize > 0 && run.FontSize < 600 && run.Text != "" {
						text := run.Text
						if len(text) > 30 {
							text = text[:30]
						}
						fmt.Printf("  Slide %d shape %d: sz=%d %q\n",
							i+1, j, run.FontSize, text)
					}
				}
			}
		}
	}

	// Issue 6: Check slide backgrounds
	fmt.Println("\n=== Slide Backgrounds ===")
	for i, s := range slides {
		bg := s.GetBackground()
		if bg.HasBackground {
			fmt.Printf("  Slide %d: fill=%q imgIdx=%d\n", i+1, bg.FillColor, bg.ImageIdx)
		}
	}

	// Issue 7: Check master backgrounds
	fmt.Println("\n=== Master Backgrounds ===")
	for ref, m := range masters {
		fmt.Printf("  Master %d: has=%v fill=%q imgIdx=%d shapes=%d scheme=%v\n",
			ref, m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx,
			len(m.Shapes), m.ColorScheme)
	}

	// Issue 8: Check for text runs with scheme color references that weren't resolved
	fmt.Println("\n=== Text Run Color Distribution (first 10 slides) ===")
	for i := 0; i < len(slides) && i < 10; i++ {
		colorDist := make(map[string]int)
		for _, sh := range slides[i].GetShapes() {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.Text == "" {
						continue
					}
					if run.Color == "" {
						colorDist["(empty)"]++
					} else {
						colorDist[run.Color]++
					}
				}
			}
		}
		if len(colorDist) > 0 {
			fmt.Printf("  Slide %d: %v\n", i+1, colorDist)
		}
	}

	// Issue 9: Check for shapes that might be watermarks (large, semi-transparent images)
	fmt.Println("\n=== Potential Watermark/Background Images ===")
	sw, sh := p.GetSlideSize()
	for ref, m := range masters {
		for i, shape := range m.Shapes {
			if shape.IsImage {
				pctW := float64(shape.Width) / float64(sw) * 100
				pctH := float64(shape.Height) / float64(sh) * 100
				fmt.Printf("  Master %d shape %d: IMAGE idx=%d size=(%d,%d) pos=(%d,%d) %.0f%%x%.0f%% opacity=%d\n",
					ref, i, shape.ImageIdx, shape.Width, shape.Height, shape.Left, shape.Top, pctW, pctH, shape.FillOpacity)
			}
		}
	}

	// Issue 10: Check for shapes with text that have no fill and no line (invisible boxes)
	fmt.Println("\n=== Text shapes with noFill+noLine (transparent text boxes) ===")
	transparentCount := 0
	for i, s := range slides {
		for _, sh := range s.GetShapes() {
			if sh.NoFill && sh.NoLine && len(sh.Paragraphs) > 0 {
				transparentCount++
				if transparentCount <= 5 {
					text := ""
					for _, p := range sh.Paragraphs {
						for _, r := range p.Runs {
							text += r.Text
						}
					}
					if len(text) > 40 {
						text = text[:40]
					}
					fmt.Printf("  Slide %d: noFill+noLine text=%q\n", i+1, text)
				}
			}
		}
	}
	fmt.Printf("  Total transparent text boxes: %d\n", transparentCount)

	_ = masters
}

func safeIdx(s []string, i int) string {
	if i < len(s) {
		return s[i]
	}
	return "(none)"
}
