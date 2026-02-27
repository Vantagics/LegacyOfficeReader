package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Parse error: %v\n", err)
		os.Exit(1)
	}

	images := p.GetImages()
	fmt.Printf("Total images: %d\n", len(images))
	// Show image 13 and 14 (0-indexed)
	for i := 12; i < 16 && i < len(images); i++ {
		img := images[i]
		fmt.Printf("  image[%d]: format=%d size=%d bytes\n", i, img.Format, len(img.Data))
	}

	// Also check slide 8 shapes for the watermark
	slides := p.GetSlides()
	if len(slides) >= 8 {
		s := slides[7]
		shapes := s.GetShapes()
		fmt.Printf("\nSlide 8 shapes: %d\n", len(shapes))
		for i, sh := range shapes {
			extra := ""
			if sh.IsImage {
				extra = fmt.Sprintf(" [IMAGE idx=%d]", sh.ImageIdx)
			}
			if sh.FillColor != "" {
				extra += fmt.Sprintf(" fill=%s", sh.FillColor)
			}
			if sh.NoFill {
				extra += " noFill"
			}
			if len(sh.GeoVertices) > 0 {
				extra += fmt.Sprintf(" [FREEFORM verts=%d]", len(sh.GeoVertices))
			}
			fmt.Printf("  shape[%d]: type=%d pos=(%d,%d) size=(%d,%d)%s\n",
				i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, extra)
		}
	}
}
