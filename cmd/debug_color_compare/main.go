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
	// Open PPT source
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	pptSlides := p.GetSlides()

	// Open PPTX output
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	issues := 0
	// For each slide, compare PPT shapes with PPTX output
	for si := 0; si < len(pptSlides) && si < 71; si++ {
		slide := pptSlides[si]
		shapes := slide.GetShapes()

		// Read PPTX slide
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si+1)
		var pptxContent string
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				pptxContent = string(data)
				break
			}
		}

		// For each PPT shape with text, check if the fill color is preserved in PPTX
		for i, sh := range shapes {
			if len(sh.Paragraphs) == 0 {
				continue
			}
			if sh.FillColor == "" || sh.NoFill {
				continue
			}

			// This shape has a fill color and text
			text := ""
			if len(sh.Paragraphs[0].Runs) > 0 {
				text = sh.Paragraphs[0].Runs[0].Text
				if len(text) > 20 {
					text = text[:20]
				}
			}

			// Check if the fill color appears in the PPTX near this text
			if text != "" && !strings.Contains(pptxContent, sh.FillColor) {
				fmt.Printf("Slide %d shape %d: PPT fill=%s MISSING in PPTX, text=%q\n", si+1, i, sh.FillColor, text)
				issues++
			}
		}
	}
	fmt.Printf("\nTotal fill color issues: %d\n", issues)
}
