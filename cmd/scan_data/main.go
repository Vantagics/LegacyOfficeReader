package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	fileName := "testfie/test.ppt"
	if len(os.Args) > 1 {
		fileName = os.Args[1]
	}

	p, err := ppt.OpenFile(fileName)
	if err != nil {
		fmt.Fprintf(os.Stderr, "open: %v\n", err)
		os.Exit(1)
	}

	fonts := p.GetFonts()
	images := p.GetImages()
	w, h := p.GetSlideSize()
	slides := p.GetSlides()

	fmt.Printf("=== PPT: %s ===\n", fileName)
	fmt.Printf("Fonts (%d): %v\n", len(fonts), fonts)
	fmt.Printf("Images: %d\n", len(images))
	fmt.Printf("Slide size: %d x %d EMU\n", w, h)
	fmt.Printf("Slides: %d\n", len(slides))

	// Per-slide stats
	emptySlides := 0
	slidesWithShapes := 0
	slidesWithImages := 0
	slidesWithFonts := 0
	slidesWithBold := 0
	totalShapes := 0
	totalImageShapes := 0
	totalTextShapes := 0

	for i, s := range slides {
		shapes := s.GetShapes()
		texts := s.GetTexts()
		hasImage := false
		hasFont := false
		hasBold := false
		imgCount := 0
		txtCount := 0

		for _, sh := range shapes {
			if sh.IsImage && sh.ImageIdx >= 0 {
				hasImage = true
				imgCount++
			}
			if sh.IsText || len(sh.Paragraphs) > 0 {
				txtCount++
			}
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.FontName != "" {
						hasFont = true
					}
					if run.Bold {
						hasBold = true
					}
				}
			}
		}

		totalShapes += len(shapes)
		totalImageShapes += imgCount
		totalTextShapes += txtCount

		if len(shapes) == 0 && len(texts) == 0 {
			emptySlides++
		}
		if len(shapes) > 0 {
			slidesWithShapes++
		}
		if hasImage {
			slidesWithImages++
		}
		if hasFont {
			slidesWithFonts++
		}
		if hasBold {
			slidesWithBold++
		}

		// Detail for first 15 slides
		if i < 15 {
			fmt.Printf("\n--- Slide %d: %d shapes, %d texts ---\n", i+1, len(shapes), len(texts))
			for j, sh := range shapes {
				if j >= 5 {
					fmt.Printf("  ... (%d more shapes)\n", len(shapes)-5)
					break
				}
				fmt.Printf("  Shape[%d]: type=%d text=%v image=%v imgIdx=%d pos=(%d,%d) size=(%dx%d)\n",
					j, sh.ShapeType, sh.IsText, sh.IsImage, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height)
				for k, para := range sh.Paragraphs {
					if k >= 3 {
						fmt.Printf("    ... (%d more paragraphs)\n", len(sh.Paragraphs)-3)
						break
					}
					for l, run := range para.Runs {
						if l >= 2 {
							break
						}
						text := run.Text
						if len(text) > 50 {
							text = text[:50] + "..."
						}
						fmt.Printf("    Para[%d] Run[%d]: font=%q sz=%d bold=%v italic=%v color=%q align=%d text=%q\n",
							k, l, run.FontName, run.FontSize, run.Bold, run.Italic, run.Color, para.Alignment, text)
					}
				}
			}
		}
	}

	fmt.Printf("\n=== PPT Summary ===\n")
	fmt.Printf("Total slides: %d\n", len(slides))
	fmt.Printf("Empty slides (no shapes, no texts): %d\n", emptySlides)
	fmt.Printf("Slides with shapes: %d\n", slidesWithShapes)
	fmt.Printf("Slides with image shapes: %d\n", slidesWithImages)
	fmt.Printf("Slides with font info: %d\n", slidesWithFonts)
	fmt.Printf("Slides with bold: %d\n", slidesWithBold)
	fmt.Printf("Total shapes: %d\n", totalShapes)
	fmt.Printf("Total image shapes: %d\n", totalImageShapes)
	fmt.Printf("Total text shapes: %d\n", totalTextShapes)
}
