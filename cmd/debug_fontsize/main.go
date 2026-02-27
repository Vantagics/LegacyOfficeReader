package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	presentation, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	slides := presentation.GetSlides()

	// Count fontSize=0 occurrences
	totalRuns := 0
	zeroRuns := 0
	zeroSlides := make(map[int]int)

	for i, s := range slides {
		shapes := s.GetShapes()
		for _, sh := range shapes {
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					totalRuns++
					if r.FontSize == 0 {
						zeroRuns++
						zeroSlides[i+1]++
					}
				}
			}
		}
	}

	fmt.Printf("Total runs: %d, fontSize=0: %d (%.1f%%)\n", totalRuns, zeroRuns, float64(zeroRuns)*100/float64(totalRuns))
	fmt.Printf("Slides with fontSize=0: %d\n", len(zeroSlides))
	for slide, count := range zeroSlides {
		fmt.Printf("  Slide %d: %d runs with fontSize=0\n", slide, count)
	}

	// Now check the PPTX output for sz="0"
	fmt.Println("\n=== PPTX sz check ===")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer zr.Close()

	sz0Count := 0
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") || strings.Contains(f.Name, "_rels") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)
		count := strings.Count(content, `sz="0"`)
		if count > 0 {
			fmt.Printf("  %s: %d occurrences of sz=\"0\"\n", f.Name, count)
			sz0Count += count
		}
	}
	if sz0Count == 0 {
		fmt.Println("  No sz=\"0\" found in any slide - GOOD")
	}

	// Check for very small font sizes that might look wrong
	fmt.Println("\n=== Small font sizes in PPTX ===")
	smallSzCount := 0
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") || strings.Contains(f.Name, "_rels") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)
		// Find sz="600" or sz="700" etc
		for _, sz := range []string{`sz="600"`, `sz="700"`, `sz="800"`, `sz="900"`} {
			count := strings.Count(content, sz)
			if count > 0 {
				smallSzCount += count
				fmt.Printf("  %s: %d occurrences of %s\n", f.Name, count, sz)
			}
		}
	}
	if smallSzCount == 0 {
		fmt.Println("  No very small font sizes found")
	}
}
