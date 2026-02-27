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

	// Check text color distribution across all slides
	colorCount := make(map[string]int)
	noColorCount := 0
	totalRuns := 0
	noFontCount := 0
	fontCount := make(map[string]int)

	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					if r.Text == "" {
						continue
					}
					totalRuns++
					if r.Color == "" {
						noColorCount++
					} else {
						colorCount[r.Color]++
					}
					if r.FontName == "" {
						noFontCount++
					} else {
						fontCount[r.FontName]++
					}
				}
			}
		}
	}

	fmt.Printf("Total text runs: %d\n", totalRuns)
	fmt.Printf("Runs with no color: %d (%.1f%%)\n", noColorCount, float64(noColorCount)*100/float64(totalRuns))
	fmt.Printf("Runs with no font: %d (%.1f%%)\n", noFontCount, float64(noFontCount)*100/float64(totalRuns))

	fmt.Println("\nTop colors:")
	type kv struct {
		k string
		v int
	}
	var sorted []kv
	for k, v := range colorCount {
		sorted = append(sorted, kv{k, v})
	}
	for i := 0; i < len(sorted); i++ {
		for j := i + 1; j < len(sorted); j++ {
			if sorted[j].v > sorted[i].v {
				sorted[i], sorted[j] = sorted[j], sorted[i]
			}
		}
	}
	for i, kv := range sorted {
		if i >= 15 {
			break
		}
		fmt.Printf("  #%s: %d runs\n", kv.k, kv.v)
	}

	fmt.Println("\nTop fonts:")
	var sortedFonts []kv
	for k, v := range fontCount {
		sortedFonts = append(sortedFonts, kv{k, v})
	}
	for i := 0; i < len(sortedFonts); i++ {
		for j := i + 1; j < len(sortedFonts); j++ {
			if sortedFonts[j].v > sortedFonts[i].v {
				sortedFonts[i], sortedFonts[j] = sortedFonts[j], sortedFonts[i]
			}
		}
	}
	for i, kv := range sortedFonts {
		if i >= 10 {
			break
		}
		fmt.Printf("  %s: %d runs\n", kv.k, kv.v)
	}

	// Check font size distribution
	fmt.Println("\nFont size distribution:")
	sizeCount := make(map[uint16]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					if r.Text != "" {
						sizeCount[r.FontSize]++
					}
				}
			}
		}
	}
	var sortedSizes []struct {
		size  uint16
		count int
	}
	for s, c := range sizeCount {
		sortedSizes = append(sortedSizes, struct {
			size  uint16
			count int
		}{s, c})
	}
	for i := 0; i < len(sortedSizes); i++ {
		for j := i + 1; j < len(sortedSizes); j++ {
			if sortedSizes[j].count > sortedSizes[i].count {
				sortedSizes[i], sortedSizes[j] = sortedSizes[j], sortedSizes[i]
			}
		}
	}
	for i, s := range sortedSizes {
		if i >= 15 {
			break
		}
		fmt.Printf("  %d (%.1fpt): %d runs\n", s.size, float64(s.size)/100, s.count)
	}
}
