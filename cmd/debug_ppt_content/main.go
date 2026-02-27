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
	fonts := p.GetFonts()
	w, h := p.GetSlideSize()
	masters := p.GetMasters()

	fmt.Printf("=== PPT Content Dump ===\n")
	fmt.Printf("Slide size: %d x %d EMU (%.1f x %.1f inches)\n", w, h, float64(w)/914400, float64(h)/914400)
	fmt.Printf("Total slides: %d\n", len(slides))
	fmt.Printf("Total images: %d\n", len(images))
	fmt.Printf("Fonts: %v\n", fonts)
	fmt.Printf("Masters: %d\n", len(masters))

	for ref, m := range masters {
		fmt.Printf("\n--- Master ref=%d ---\n", ref)
		fmt.Printf("  Background: has=%v color=%s imgIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
		fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
		fmt.Printf("  Shapes: %d\n", len(m.Shapes))
		for si, sh := range m.Shapes {
			fmt.Printf("  Shape[%d]: type=%d pos=(%d,%d) size=(%d,%d) isText=%v isImage=%v imgIdx=%d\n",
				si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsText, sh.IsImage, sh.ImageIdx)
			fmt.Printf("    fill=%s noFill=%v line=%s noLine=%v lineW=%d rot=%d flipH=%v flipV=%v\n",
				sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine, sh.LineWidth, sh.Rotation, sh.FlipH, sh.FlipV)
			fmt.Printf("    textMargins: L=%d T=%d R=%d B=%d anchor=%d wrap=%d opacity=%d\n",
				sh.TextMarginLeft, sh.TextMarginTop, sh.TextMarginRight, sh.TextMarginBottom, sh.TextAnchor, sh.TextWordWrap, sh.FillOpacity)
			for pi, para := range sh.Paragraphs {
				fmt.Printf("    Para[%d]: align=%d indent=%d bullet=%v bulletChar=%q lm=%d ind=%d\n",
					pi, para.Alignment, para.IndentLevel, para.HasBullet, para.BulletChar, para.LeftMargin, para.Indent)
				fmt.Printf("      spacing: before=%d after=%d line=%d\n", para.SpaceBefore, para.SpaceAfter, para.LineSpacing)
				for ri, run := range para.Runs {
					text := run.Text
					if len(text) > 80 {
						text = text[:80] + "..."
					}
					text = strings.ReplaceAll(text, "\n", "\\n")
					text = strings.ReplaceAll(text, "\r", "\\r")
					fmt.Printf("      Run[%d]: font=%q size=%d bold=%v italic=%v underline=%v color=%s text=%q\n",
						ri, run.FontName, run.FontSize, run.Bold, run.Italic, run.Underline, run.Color, text)
				}
			}
		}
	}

	for i, slide := range slides {
		fmt.Printf("\n=== Slide %d ===\n", i+1)
		fmt.Printf("Layout type: %d\n", slide.GetLayoutType())
		fmt.Printf("Master ref: %d\n", slide.GetMasterRef())
		bg := slide.GetBackground()
		fmt.Printf("Background: has=%v color=%s imgIdx=%d\n", bg.HasBackground, bg.FillColor, bg.ImageIdx)

		shapes := slide.GetShapes()
		fmt.Printf("Shapes: %d\n", len(shapes))
		for si, sh := range shapes {
			fmt.Printf("  Shape[%d]: type=%d pos=(%d,%d) size=(%d,%d) isText=%v isImage=%v imgIdx=%d\n",
				si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsText, sh.IsImage, sh.ImageIdx)
			fmt.Printf("    fill=%s noFill=%v line=%s noLine=%v lineW=%d rot=%d flipH=%v flipV=%v\n",
				sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine, sh.LineWidth, sh.Rotation, sh.FlipH, sh.FlipV)
			fmt.Printf("    textMargins: L=%d T=%d R=%d B=%d anchor=%d wrap=%d opacity=%d\n",
				sh.TextMarginLeft, sh.TextMarginTop, sh.TextMarginRight, sh.TextMarginBottom, sh.TextAnchor, sh.TextWordWrap, sh.FillOpacity)
			for pi, para := range sh.Paragraphs {
				fmt.Printf("    Para[%d]: align=%d indent=%d bullet=%v bulletChar=%q lm=%d ind=%d\n",
					pi, para.Alignment, para.IndentLevel, para.HasBullet, para.BulletChar, para.LeftMargin, para.Indent)
				fmt.Printf("      spacing: before=%d after=%d line=%d\n", para.SpaceBefore, para.SpaceAfter, para.LineSpacing)
				for ri, run := range para.Runs {
					text := run.Text
					if len(text) > 120 {
						text = text[:120] + "..."
					}
					text = strings.ReplaceAll(text, "\n", "\\n")
					text = strings.ReplaceAll(text, "\r", "\\r")
					text = strings.ReplaceAll(text, "\x0b", "\\v")
					fmt.Printf("      Run[%d]: font=%q size=%d bold=%v italic=%v underline=%v color=%s text=%q\n",
						ri, run.FontName, run.FontSize, run.Bold, run.Italic, run.Underline, run.Color, text)
				}
			}
		}

		texts := slide.GetTexts()
		if len(texts) > 0 {
			fmt.Printf("  Raw texts (%d):\n", len(texts))
			for ti, t := range texts {
				t2 := strings.ReplaceAll(t, "\n", "\\n")
				t2 = strings.ReplaceAll(t2, "\r", "\\r")
				t2 = strings.ReplaceAll(t2, "\x0b", "\\v")
				if len(t2) > 120 {
					t2 = t2[:120] + "..."
				}
				fmt.Printf("    [%d]: %q\n", ti, t2)
			}
		}
	}

	for i, img := range images {
		fmt.Printf("\nImage[%d]: format=%d size=%d bytes\n", i, img.Format, len(img.Data))
	}
}
