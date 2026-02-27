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

	// Check color scheme usage per slide
	for i, s := range slides {
		ref := s.GetMasterRef()
		m, ok := masters[ref]
		if !ok {
			continue
		}
		scheme := m.ColorScheme
		if len(scheme) < 2 {
			continue
		}
		textColor := scheme[1] // dk1 = text color

		shapes := s.GetShapes()
		emptyColorCount := 0
		totalRuns := 0
		for _, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					totalRuns++
					if run.Color == "" {
						emptyColorCount++
					}
				}
			}
		}
		if emptyColorCount > 0 {
			fmt.Printf("Slide %d: master=%d scheme_text=%s emptyColors=%d/%d\n",
				i+1, ref, textColor, emptyColorCount, totalRuns)
			// Show first empty color run
			for _, sh := range shapes {
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						if run.Color == "" {
							text := run.Text
							if len(text) > 40 {
								text = text[:40] + "..."
							}
							text = strings.ReplaceAll(text, "\n", "\\n")
							text = strings.ReplaceAll(text, "\x0b", "\\v")
							fmt.Printf("  empty color: font=%q sz=%d text=%q\n", run.FontName, run.FontSize, text)
							goto nextSlide
						}
					}
				}
			}
		}
	nextSlide:
	}
}
