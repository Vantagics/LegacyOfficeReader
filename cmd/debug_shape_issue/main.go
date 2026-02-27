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
	fmt.Printf("Slide size: %d x %d EMU\n", w, h)

	// Find slides with shapes that look like the screenshot:
	// - Has images (blue shapes with icons)
	// - Has text about "核心能力" or "资产发现" etc.
	// - Has watermark at bottom ("奇安信" or "中国奥委会")
	for i, slide := range slides {
		shapes := slide.GetShapes()
		texts := slide.GetTexts()

		// Check all text content
		allText := ""
		for _, t := range texts {
			allText += t
		}
		for _, sh := range shapes {
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					allText += r.Text
				}
			}
		}

		// Look for slides with group shapes or many shapes (the screenshot has complex layout)
		hasGroupShapes := false
		hasImages := false
		hasWatermark := false
		for _, sh := range shapes {
			if sh.IsImage {
				hasImages = true
			}
			if sh.ShapeType == 0 {
				hasGroupShapes = true
			}
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					if strings.Contains(r.Text, "qianxin") || strings.Contains(r.Text, "www.") {
						hasWatermark = true
					}
				}
			}
		}

		// Print slides with interesting content
		if len(shapes) > 3 || hasGroupShapes || hasImages {
			fmt.Printf("\n=== Slide %d (shapes=%d, texts=%d, hasGroup=%v, hasImages=%v, hasWatermark=%v) ===\n",
				i+1, len(shapes), len(texts), hasGroupShapes, hasImages, hasWatermark)
			fmt.Printf("Layout: %d, MasterRef: %d\n", slide.GetLayoutType(), slide.GetMasterRef())
			bg := slide.GetBackground()
			fmt.Printf("Background: has=%v color=%s imgIdx=%d\n", bg.HasBackground, bg.FillColor, bg.ImageIdx)

			for si, sh := range shapes {
				fmt.Printf("  Shape[%d]: type=%d pos=(%d,%d) size=(%d,%d) isText=%v isImage=%v imgIdx=%d\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsText, sh.IsImage, sh.ImageIdx)
				fmt.Printf("    fill=%s noFill=%v line=%s noLine=%v lineW=%d rot=%d flipH=%v flipV=%v opacity=%d\n",
					sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine, sh.LineWidth, sh.Rotation, sh.FlipH, sh.FlipV, sh.FillOpacity)
				for pi, para := range sh.Paragraphs {
					for ri, run := range para.Runs {
						text := run.Text
						if len(text) > 100 {
							text = text[:100] + "..."
						}
						fmt.Printf("    Para[%d].Run[%d]: font=%q size=%d color=%s colorRaw=0x%08X text=%q\n",
							pi, ri, run.FontName, run.FontSize, run.Color, run.ColorRaw, text)
					}
				}
			}

			// Also show raw texts
			for ti, t := range texts {
				t2 := t
				if len(t2) > 100 {
					t2 = t2[:100] + "..."
				}
				t2 = strings.ReplaceAll(t2, "\r", "\\r")
				t2 = strings.ReplaceAll(t2, "\n", "\\n")
				fmt.Printf("  Text[%d]: %q\n", ti, t2)
			}
		}
	}

	// Print master info
	for ref, m := range masters {
		fmt.Printf("\n--- Master ref=%d ---\n", ref)
		fmt.Printf("  Background: has=%v color=%s imgIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
		fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
		fmt.Printf("  Shapes: %d\n", len(m.Shapes))
		for si, sh := range m.Shapes {
			fmt.Printf("  MasterShape[%d]: type=%d pos=(%d,%d) size=(%d,%d) isText=%v isImage=%v imgIdx=%d\n",
				si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsText, sh.IsImage, sh.ImageIdx)
			fmt.Printf("    fill=%s noFill=%v line=%s noLine=%v lineW=%d rot=%d flipH=%v flipV=%v opacity=%d\n",
				sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine, sh.LineWidth, sh.Rotation, sh.FlipH, sh.FlipV, sh.FillOpacity)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					text := run.Text
					if len(text) > 100 {
						text = text[:100] + "..."
					}
					fmt.Printf("    Para[%d].Run[%d]: font=%q size=%d color=%s text=%q\n",
						pi, ri, run.FontName, run.FontSize, run.Color, text)
				}
			}
		}
	}
}
