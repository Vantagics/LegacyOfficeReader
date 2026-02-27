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
	// Parse PPT
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPT: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	fmt.Printf("PPT: %d slides\n", len(slides))

	// Count PPTX slides
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	pptxSlideCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			pptxSlideCount++
		}
	}
	fmt.Printf("PPTX: %d slides\n", pptxSlideCount)

	// Dump first slide XML from PPTX
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if len(content) > 3000 {
				content = content[:3000] + "..."
			}
			fmt.Printf("\n=== PPTX slide1.xml (first 3000 chars) ===\n%s\n", content)
			break
		}
	}

	// Show PPT slide 1 content
	if len(slides) > 0 {
		fmt.Printf("\n=== PPT Slide 1 ===\n")
		s := slides[0]
		fmt.Printf("Layout: %d, Master: %d\n", s.GetLayoutType(), s.GetMasterRef())
		shapes := s.GetShapes()
		fmt.Printf("Shapes: %d\n", len(shapes))
		for i, sh := range shapes {
			fmt.Printf("  [%d] type=%d pos=(%d,%d) size=(%d,%d) img=%v text=%v\n",
				i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsImage, sh.IsText)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					text := run.Text
					if len(text) > 60 {
						text = text[:60] + "..."
					}
					text = strings.ReplaceAll(text, "\n", "\\n")
					text = strings.ReplaceAll(text, "\r", "\\r")
					fmt.Printf("    P%d.R%d: font=%q sz=%d color=%s text=%q\n",
						pi, ri, run.FontName, run.FontSize, run.Color, text)
				}
			}
		}
	}

	// Show layout XML from PPTX
	for _, f := range zr.File {
		if f.Name == "ppt/slideLayouts/slideLayout1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if len(content) > 2000 {
				content = content[:2000] + "..."
			}
			fmt.Printf("\n=== PPTX slideLayout1.xml (first 2000 chars) ===\n%s\n", content)
			break
		}
	}
}
