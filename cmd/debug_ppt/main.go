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

	fmt.Printf("=== PPT Analysis ===\n")
	fmt.Printf("Fonts: %v\n", p.GetFonts())
	w, h := p.GetSlideSize()
	fmt.Printf("Slide size: %d x %d EMU\n", w, h)
	fmt.Printf("Images: %d\n", len(p.GetImages()))
	for i, img := range p.GetImages() {
		fmt.Printf("  Image %d: format=%d, size=%d bytes\n", i, img.Format, len(img.Data))
	}
	fmt.Printf("Slides: %d\n", p.GetNumberSlides())

	for i, slide := range p.GetSlides() {
		fmt.Printf("\n--- Slide %d (layout=%d, masterRef=%d) ---\n", i+1, slide.GetLayoutType(), slide.GetMasterRef())
		bg := slide.GetBackground()
		fmt.Printf("Background: has=%v, color=%s, imgIdx=%d\n", bg.HasBackground, bg.FillColor, bg.ImageIdx)

		texts := slide.GetTexts()
		fmt.Printf("Texts (%d):\n", len(texts))
		for j, t := range texts {
			fmt.Printf("  [%d] %q\n", j, truncate(t, 100))
		}

		shapes := slide.GetShapes()
		fmt.Printf("Shapes (%d):\n", len(shapes))
		for j, sh := range shapes {
			fmt.Printf("  Shape %d: type=%d, pos=(%d,%d), size=(%d,%d), isText=%v, isImage=%v, imgIdx=%d\n",
				j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsText, sh.IsImage, sh.ImageIdx)
			if sh.FillColor != "" {
				fmt.Printf("    fill=%s", sh.FillColor)
			}
			if sh.LineColor != "" {
				fmt.Printf("    line=%s", sh.LineColor)
			}
			if sh.NoFill {
				fmt.Printf("    noFill")
			}
			if sh.NoLine {
				fmt.Printf("    noLine")
			}
			if sh.Rotation != 0 {
				fmt.Printf("    rot=%d", sh.Rotation)
			}
			if sh.FlipH {
				fmt.Printf("    flipH")
			}
			if sh.FlipV {
				fmt.Printf("    flipV")
			}
			fmt.Println()
			for k, para := range sh.Paragraphs {
				fmt.Printf("    Para %d: align=%d, indent=%d, bullet=%v, bulletChar=%q, lnSpc=%d, spcBef=%d, spcAft=%d, marL=%d, indent=%d\n",
					k, para.Alignment, para.IndentLevel, para.HasBullet, para.BulletChar, para.LineSpacing, para.SpaceBefore, para.SpaceAfter, para.LeftMargin, para.Indent)
				for l, run := range para.Runs {
					fmt.Printf("      Run %d: font=%q, size=%d, bold=%v, italic=%v, underline=%v, color=%s, text=%q\n",
						l, run.FontName, run.FontSize, run.Bold, run.Italic, run.Underline, run.Color, truncate(run.Text, 80))
				}
			}
		}
	}
}

func truncate(s string, n int) string {
	s = strings.ReplaceAll(s, "\n", "\\n")
	s = strings.ReplaceAll(s, "\r", "\\r")
	if len(s) > n {
		return s[:n] + "..."
	}
	return s
}
