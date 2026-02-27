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
	masters := p.GetMasters()
	w, h := p.GetSlideSize()
	fmt.Printf("Slide size: %d x %d\n", w, h)

	// Check which masters have watermark-like images
	fmt.Println("\n=== Master watermark analysis ===")
	for ref, m := range masters {
		for si, sh := range m.Shapes {
			if sh.IsImage && sh.ImageIdx >= 0 {
				bottom := int64(sh.Top) + int64(sh.Height)
				isWatermark := int64(sh.Top) > int64(h)/2 && bottom > int64(h)*3/4
				isFullPage := sh.Width > int32(float64(w)*0.7) && sh.Height > int32(float64(h)*0.7)
				fmt.Printf("  Master %d Shape[%d]: imgIdx=%d pos=(%d,%d) sz=(%d,%d) fullPage=%v watermark=%v bottom=%d slideH=%d\n",
					ref, si, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height, isFullPage, isWatermark, bottom, h)
			}
		}
	}

	// Check text color issues on specific slides
	fmt.Println("\n=== Text color analysis (slides with potential issues) ===")
	for i, s := range slides {
		shapes := s.GetShapes()
		ref := s.GetMasterRef()
		m, hasMaster := masters[ref]
		var colorScheme []string
		if hasMaster {
			colorScheme = m.ColorScheme
		}

		hasColorIssue := false
		for _, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					// Check for white text on shapes without dark fill
					if run.Color == "FFFFFF" && sh.FillColor == "" && sh.NoFill {
						hasColorIssue = true
					}
					// Check for empty color
					if run.Color == "" && run.Text != "" && strings.TrimSpace(run.Text) != "" {
						hasColorIssue = true
					}
				}
			}
		}

		if hasColorIssue && i < 20 {
			fmt.Printf("\nSlide %d (masterRef=%d, scheme=%v):\n", i+1, ref, colorScheme)
			for si, sh := range shapes {
				if !sh.IsText || len(sh.Paragraphs) == 0 {
					continue
				}
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						if (run.Color == "FFFFFF" && sh.FillColor == "" && sh.NoFill) ||
							(run.Color == "" && strings.TrimSpace(run.Text) != "") {
							text := run.Text
							if len(text) > 60 {
								text = text[:60] + "..."
							}
							text = strings.ReplaceAll(text, "\n", "\\n")
							text = strings.ReplaceAll(text, "\r", "\\r")
							text = strings.ReplaceAll(text, "\x0b", "\\v")
							fmt.Printf("  Shape[%d] fill=%s noFill=%v: color=%s colorRaw=0x%08X text=%q\n",
								si, sh.FillColor, sh.NoFill, run.Color, run.ColorRaw, text)
						}
					}
				}
			}
		}
	}

	// Check slide 4 specifically (E9EBF5 fill with white text issue)
	fmt.Println("\n=== Slide 4 detailed color check ===")
	if len(slides) >= 4 {
		s := slides[3]
		shapes := s.GetShapes()
		for si, sh := range shapes {
			if !sh.IsText || len(sh.Paragraphs) == 0 {
				continue
			}
			if si > 20 {
				break
			}
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					text := run.Text
					if len(text) > 50 {
						text = text[:50] + "..."
					}
					text = strings.ReplaceAll(text, "\x0b", "\\v")
					fmt.Printf("  S4 Shape[%d] P[%d] R[%d]: fill=%s noFill=%v color=%s raw=0x%08X sz=%d b=%v text=%q\n",
						si, pi, ri, sh.FillColor, sh.NoFill, run.Color, run.ColorRaw, run.FontSize, run.Bold, text)
				}
			}
		}
	}

	// Check slide 2 text colors (CONTENTS text should be DDDDDD)
	fmt.Println("\n=== Slide 2 text colors ===")
	if len(slides) >= 2 {
		s := slides[1]
		shapes := s.GetShapes()
		for si, sh := range shapes {
			if !sh.IsText || len(sh.Paragraphs) == 0 {
				continue
			}
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					text := run.Text
					if len(text) > 50 {
						text = text[:50] + "..."
					}
					fmt.Printf("  S2 Shape[%d]: fill=%s noFill=%v color=%s raw=0x%08X text=%q\n",
						si, sh.FillColor, sh.NoFill, run.Color, run.ColorRaw, text)
				}
			}
		}
	}
}
