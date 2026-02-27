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
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	// Check slides with biggest mismatches: 4, 9, 26, 63, 69
	checkSlides := []int{3, 8, 25, 62, 68} // 0-indexed

	for _, idx := range checkSlides {
		if idx >= len(slides) {
			continue
		}
		s := slides[idx]
		shapes := s.GetShapes()
		fmt.Printf("\n=== Slide %d (%d shapes) ===\n", idx+1, len(shapes))

		// Count by type
		typeCounts := make(map[string]int)
		noFillNoLineCount := 0
		emptyTextCount := 0
		zeroSizeCount := 0

		for i, sh := range shapes {
			key := fmt.Sprintf("type=%d", sh.ShapeType)
			if sh.IsImage {
				key = "image"
			} else if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
				key = "connector"
			} else if sh.IsText {
				key = "textbox"
			}
			typeCounts[key]++

			// Check for shapes that might be filtered
			if sh.NoFill && sh.NoLine && !sh.IsImage && len(sh.Paragraphs) == 0 {
				noFillNoLineCount++
			}

			// Check for empty text shapes
			hasText := false
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if strings.TrimSpace(run.Text) != "" {
						hasText = true
						break
					}
				}
				if hasText {
					break
				}
			}
			if !sh.IsImage && !hasText && sh.ShapeType != 20 && !(sh.ShapeType >= 32 && sh.ShapeType <= 40) {
				emptyTextCount++
			}

			// Check for zero-size shapes
			if sh.Width == 0 && sh.Height == 0 {
				zeroSizeCount++
			}

			// Print first 10 shapes for detail
			if i < 20 || i >= len(shapes)-5 {
				textSnippet := ""
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						t := strings.TrimSpace(run.Text)
						if t != "" {
							if len(t) > 30 {
								textSnippet = t[:30]
							} else {
								textSnippet = t
							}
							break
						}
					}
					if textSnippet != "" {
						break
					}
				}
				fmt.Printf("  [%3d] type=%3d img=%v txt=%v noFill=%v noLine=%v w=%d h=%d paras=%d  %s\n",
					i, sh.ShapeType, sh.IsImage, sh.IsText, sh.NoFill, sh.NoLine,
					sh.Width, sh.Height, len(sh.Paragraphs), textSnippet)
			}
		}

		fmt.Printf("  Type distribution: %v\n", typeCounts)
		fmt.Printf("  NoFill+NoLine+NoImage+NoParagraphs: %d\n", noFillNoLineCount)
		fmt.Printf("  Empty text shapes: %d\n", emptyTextCount)
		fmt.Printf("  Zero-size shapes: %d\n", zeroSizeCount)
	}
}
