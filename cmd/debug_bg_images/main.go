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

	images := p.GetImages()
	masters := p.GetMasters()
	slides := p.GetSlides()

	// Show which images are used as backgrounds
	fmt.Println("=== Master background images ===")
	for ref, m := range masters {
		if m.Background.ImageIdx >= 0 {
			img := images[m.Background.ImageIdx]
			fmt.Printf("Master %d: bgImg=%d format=%d size=%d bytes\n",
				ref, m.Background.ImageIdx, img.Format, len(img.Data))
		}
		// Check master shapes for full-page images
		for _, sh := range m.Shapes {
			if sh.IsImage && sh.ImageIdx >= 0 {
				img := images[sh.ImageIdx]
				fmt.Printf("Master %d: shapeImg=%d format=%d size=%d pos=(%d,%d) size=(%d,%d)\n",
					ref, sh.ImageIdx, img.Format, len(img.Data), sh.Left, sh.Top, sh.Width, sh.Height)
			}
		}
	}

	fmt.Println("\n=== Slide background images ===")
	for i, s := range slides {
		bg := s.GetBackground()
		if bg.ImageIdx >= 0 {
			img := images[bg.ImageIdx]
			fmt.Printf("Slide %d: bgImg=%d format=%d size=%d bytes\n",
				i+1, bg.ImageIdx, img.Format, len(img.Data))
		}
		// Check slide shapes for images
		imgCount := 0
		for _, sh := range s.GetShapes() {
			if sh.IsImage {
				imgCount++
			}
		}
		if imgCount > 0 {
			fmt.Printf("Slide %d: %d image shapes\n", i+1, imgCount)
		}
	}
}
