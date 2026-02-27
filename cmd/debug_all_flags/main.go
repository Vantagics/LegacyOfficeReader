package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		panic(err)
	}

	slides := p.GetSlides()
	// Count all unique flag values across all text runs
	flagCounts := map[uint32]int{}
	for _, s := range slides {
		shapes := s.GetShapes()
		for _, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.ColorRaw != 0 {
						flag := run.ColorRaw >> 24
						flagCounts[flag]++
					}
				}
			}
		}
	}
	fmt.Println("Text run color flags:")
	for flag, count := range flagCounts {
		fmt.Printf("  0x%02X: %d runs\n", flag, count)
	}

	// Also check fill color flags
	fillFlagCounts := map[uint32]int{}
	for _, s := range slides {
		shapes := s.GetShapes()
		for _, sh := range shapes {
			if sh.FillColorRaw != 0 {
				flag := sh.FillColorRaw >> 24
				fillFlagCounts[flag]++
			}
		}
	}
	fmt.Println("\nFill color flags:")
	for flag, count := range fillFlagCounts {
		fmt.Printf("  0x%02X: %d shapes\n", flag, count)
	}
}
