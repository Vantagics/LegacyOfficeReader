package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"

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

	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		panic(err)
	}
	defer zr.Close()

	// For each slide, compare PPT shapes with PPTX content
	for i := 0; i < len(slides) && i < 5; i++ {
		s := slides[i]
		shapes := s.GetShapes()
		cs := s.GetColorScheme()

		fmt.Printf("\n========== Slide %d ==========\n", i+1)
		fmt.Printf("PPT: %d shapes, scheme=%v\n", len(shapes), cs)

		// Count text runs and their properties
		totalRuns := 0
		noColor := 0
		noFont := 0
		noSize := 0
		for _, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.Text == "" {
						continue
					}
					totalRuns++
					if run.Color == "" {
						noColor++
					}
					if run.FontName == "" {
						noFont++
					}
					if run.FontSize == 0 {
						noSize++
					}
				}
			}
		}
		fmt.Printf("PPT runs: total=%d noColor=%d noFont=%d noSize=%d\n", totalRuns, noColor, noFont, noSize)

		// Check PPTX output
		name := fmt.Sprintf("ppt/slides/slide%d.xml", i+1)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				pptxRuns := strings.Count(content, "<a:r>")
				pptxBr := strings.Count(content, "<a:br>")
				hasSz0 := strings.Contains(content, `sz="0"`)
				hasEmptyVal := strings.Contains(content, `val=""`)

				fmt.Printf("PPTX: runs=%d breaks=%d sz0=%v emptyVal=%v size=%d bytes\n",
					pptxRuns, pptxBr, hasSz0, hasEmptyVal, len(data))
			}
		}
	}

	// Check specific problematic slides
	fmt.Println("\n========== Checking all slides for issues ==========")
	for i := 0; i < len(slides); i++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", i+1)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				issues := []string{}
				if strings.Contains(content, `sz="0"`) {
					issues = append(issues, "sz=0")
				}
				if strings.Contains(content, `val=""/>`) {
					issues = append(issues, "empty-val")
				}
				if strings.Contains(content, `typeface=""/>`) {
					issues = append(issues, "empty-typeface")
				}
				// Check for very large font sizes (> 6000 = 60pt)
				// This could indicate estimation errors
				if strings.Contains(content, `sz="7200"`) || strings.Contains(content, `sz="8000"`) || strings.Contains(content, `sz="9600"`) {
					issues = append(issues, "very-large-font")
				}

				if len(issues) > 0 {
					fmt.Printf("  Slide %d: %v\n", i+1, issues)
				}
			}
		}
	}
	fmt.Println("Done.")
}
