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

	// Count shapes by fill status
	noFillExplicit := 0
	hasFillColor := 0
	noFillNoColor := 0 // neither noFill nor fillColor - these might get default fill
	totalShapes := 0

	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			totalShapes++
			if sh.NoFill {
				noFillExplicit++
			} else if sh.FillColor != "" {
				hasFillColor++
			} else {
				noFillNoColor++
			}
		}
	}

	fmt.Printf("=== Fill Status ===\n")
	fmt.Printf("Total shapes: %d\n", totalShapes)
	fmt.Printf("Explicit noFill: %d\n", noFillExplicit)
	fmt.Printf("Has fill color: %d\n", hasFillColor)
	fmt.Printf("No fill, no color (ambiguous): %d\n", noFillNoColor)

	// Check what types these ambiguous shapes are
	fmt.Printf("\n=== Ambiguous Fill Shapes by Type ===\n")
	typeDist := make(map[uint16]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			if !sh.NoFill && sh.FillColor == "" {
				typeDist[sh.ShapeType]++
			}
		}
	}
	for t, count := range typeDist {
		fmt.Printf("  type=%d: %d\n", t, count)
	}

	// Check line status for ambiguous shapes
	fmt.Printf("\n=== Ambiguous Fill Shapes - Line Status ===\n")
	noLineExplicit := 0
	hasLineColor := 0
	noLineNoColor := 0
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			if !sh.NoFill && sh.FillColor == "" {
				if sh.NoLine {
					noLineExplicit++
				} else if sh.LineColor != "" {
					hasLineColor++
				} else {
					noLineNoColor++
				}
			}
		}
	}
	fmt.Printf("  Explicit noLine: %d\n", noLineExplicit)
	fmt.Printf("  Has line color: %d\n", hasLineColor)
	fmt.Printf("  No line, no color: %d\n", noLineNoColor)

	// Show first 10 ambiguous shapes
	fmt.Printf("\n=== First 10 Ambiguous Fill Shapes ===\n")
	count := 0
	for si, s := range slides {
		for i, sh := range s.GetShapes() {
			if !sh.NoFill && sh.FillColor == "" && !sh.IsImage {
				textSnippet := ""
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						t := strings.TrimSpace(run.Text)
						if t != "" {
							if len([]rune(t)) > 20 {
								textSnippet = string([]rune(t)[:20])
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
				fmt.Printf("  Slide %d [%d] type=%d noLine=%v lineColor=%s w=%d h=%d text=%s\n",
					si+1, i, sh.ShapeType, sh.NoLine, sh.LineColor, sh.Width, sh.Height, textSnippet)
				count++
				if count >= 10 {
					return
				}
			}
		}
	}
}
