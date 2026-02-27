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
	w, h := p.GetSlideSize()

	fmt.Printf("Slides: %d, Size: %dx%d EMU\n", len(slides), w, h)

	for i, slide := range slides {
		fmt.Printf("\n=== SLIDE %d (layout=%d master=%d) ===\n", i+1, slide.GetLayoutType(), slide.GetMasterRef())
		bg := slide.GetBackground()
		if bg.HasBackground {
			fmt.Printf("BG: color=%s imgIdx=%d\n", bg.FillColor, bg.ImageIdx)
		}
		shapes := slide.GetShapes()
		fmt.Printf("Shapes: %d\n", len(shapes))
		for si, sh := range shapes {
			fmt.Printf("  [%d] type=%d (%dx%d @ %d,%d) img=%v(idx=%d) text=%v\n",
				si, sh.ShapeType, sh.Width, sh.Height, sh.Left, sh.Top, sh.IsImage, sh.ImageIdx, sh.IsText)
			if sh.FillColor != "" || sh.NoFill {
				fmt.Printf("       fill=%s noFill=%v opacity=%d\n", sh.FillColor, sh.NoFill, sh.FillOpacity)
			}
			if sh.LineColor != "" || sh.NoLine {
				fmt.Printf("       line=%s noLine=%v lineW=%d\n", sh.LineColor, sh.NoLine, sh.LineWidth)
			}
			if sh.Rotation != 0 || sh.FlipH || sh.FlipV {
				fmt.Printf("       rot=%d flipH=%v flipV=%v\n", sh.Rotation, sh.FlipH, sh.FlipV)
			}
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					text := strings.ReplaceAll(run.Text, "\n", "\\n")
					text = strings.ReplaceAll(text, "\r", "\\r")
					text = strings.ReplaceAll(text, "\x0b", "\\v")
					if len(text) > 100 {
						text = text[:100] + "..."
					}
					fmt.Printf("       P%d.R%d: align=%d font=%q sz=%d b=%v i=%v u=%v color=%s text=%q\n",
						pi, ri, para.Alignment, run.FontName, run.FontSize, run.Bold, run.Italic, run.Underline, run.Color, text)
					if para.HasBullet {
						fmt.Printf("              bullet=%q bulletColor=%s bulletFont=%q\n", para.BulletChar, para.BulletColor, para.BulletFont)
					}
				}
			}
		}
	}
}
