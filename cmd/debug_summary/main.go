package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	images := p.GetImages()
	w, h := p.GetSlideSize()
	masters := p.GetMasters()

	fmt.Printf("Slide size: %d x %d EMU\n", w, h)
	fmt.Printf("Slides: %d, Images: %d, Masters: %d\n\n", len(slides), len(images), len(masters))

	// Dump masters
	for ref, m := range masters {
		fmt.Printf("Master ref=%d: bg.has=%v bg.color=%s bg.imgIdx=%d scheme=%v shapes=%d\n",
			ref, m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx, m.ColorScheme, len(m.Shapes))
		for si, sh := range m.Shapes {
			if si > 5 {
				fmt.Printf("  ... and %d more shapes\n", len(m.Shapes)-6)
				break
			}
			desc := fmt.Sprintf("type=%d %dx%d", sh.ShapeType, sh.Width, sh.Height)
			if sh.IsImage {
				desc += fmt.Sprintf(" IMG[%d]", sh.ImageIdx)
			}
			if sh.IsText && len(sh.Paragraphs) > 0 {
				var texts []string
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						t := strings.TrimSpace(run.Text)
						if t != "" {
							if len(t) > 40 {
								t = t[:40] + "..."
							}
							texts = append(texts, t)
						}
					}
				}
				if len(texts) > 0 {
					desc += " TEXT:" + strings.Join(texts, "|")
				}
			}
			fmt.Printf("  Shape[%d]: %s\n", si, desc)
		}
		// Show default text styles
		for lvl, style := range m.DefaultTextStyles {
			if style.FontSize > 0 || style.FontName != "" || style.Color != "" {
				fmt.Printf("  DefTextStyle[%d]: size=%d font=%q color=%s bold=%v\n",
					lvl, style.FontSize, style.FontName, style.Color, style.Bold)
			}
		}
	}

	// Dump first 10 slides and last 5 slides
	dumpSlide := func(i int) {
		slide := slides[i]
		bg := slide.GetBackground()
		shapes := slide.GetShapes()
		fmt.Printf("\n--- Slide %d (master=%d) ---\n", i+1, slide.GetMasterRef())
		fmt.Printf("  bg: has=%v color=%s imgIdx=%d\n", bg.HasBackground, bg.FillColor, bg.ImageIdx)
		fmt.Printf("  shapes: %d\n", len(shapes))
		for si, sh := range shapes {
			if si > 8 {
				fmt.Printf("  ... and %d more shapes\n", len(shapes)-9)
				break
			}
			desc := fmt.Sprintf("type=%d pos=(%d,%d) %dx%d", sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height)
			if sh.IsImage {
				desc += fmt.Sprintf(" IMG[%d]", sh.ImageIdx)
			}
			if sh.FillColor != "" {
				desc += " fill=" + sh.FillColor
			}
			if sh.NoFill {
				desc += " noFill"
			}
			if sh.IsText && len(sh.Paragraphs) > 0 {
				for pi, para := range sh.Paragraphs {
					if pi > 3 {
						fmt.Printf("    ... and %d more paragraphs\n", len(sh.Paragraphs)-4)
						break
					}
					for ri, run := range para.Runs {
						if ri > 2 {
							break
						}
						t := strings.TrimSpace(run.Text)
						if len(t) > 60 {
							t = t[:60] + "..."
						}
						t = strings.ReplaceAll(t, "\n", "\\n")
						t = strings.ReplaceAll(t, "\x0b", "\\v")
						fmt.Printf("    P[%d]R[%d]: font=%q sz=%d color=%s bold=%v text=%q\n",
							pi, ri, run.FontName, run.FontSize, run.Color, run.Bold, t)
					}
				}
			}
			fmt.Printf("  S[%d]: %s\n", si, desc)
		}
	}

	for i := 0; i < len(slides) && i < 10; i++ {
		dumpSlide(i)
	}
	if len(slides) > 10 {
		fmt.Printf("\n... skipping slides 11-%d ...\n", len(slides)-5)
		for i := len(slides) - 5; i < len(slides); i++ {
			dumpSlide(i)
		}
	}
}
