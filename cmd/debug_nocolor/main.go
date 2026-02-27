package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	// Check slides with no-color text runs and their master backgrounds
	fmt.Println("Slides with no-color text runs:")
	for i, s := range slides {
		hasNoColor := false
		for _, sh := range s.GetShapes() {
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					if r.Text != "" && r.Color == "" {
						hasNoColor = true
						break
					}
				}
				if hasNoColor {
					break
				}
			}
			if hasNoColor {
				break
			}
		}
		if !hasNoColor {
			continue
		}

		// Show the no-color runs with their shape context
		for j, sh := range s.GetShapes() {
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					if r.Text != "" && r.Color == "" {
						text := r.Text
						if len(text) > 50 {
							text = text[:50] + "..."
						}
						fmt.Printf("  Slide %d, Shape %d (fill=%s, noFill=%v): font=%s, size=%d, text=%q\n",
							i+1, j, sh.FillColor, sh.NoFill, r.FontName, r.FontSize, text)
						break // just show first no-color run per shape
					}
				}
			}
		}
	}
}
