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
	// Check what images are on slide 71
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	slide := slides[70]
	shapes := slide.GetShapes()
	images := p.GetImages()

	fmt.Printf("Total images in PPT: %d\n", len(images))

	for i, sh := range shapes {
		if sh.IsImage {
			fmt.Printf("\nShape[%d]: IMG[%d] pos=(%d,%d) sz=(%d,%d)\n",
				i, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height)
			if sh.ImageIdx >= 0 && sh.ImageIdx < len(images) {
				img := images[sh.ImageIdx]
				fmt.Printf("  Image format: %d (%s), size: %d bytes\n", img.Format, img.Extension(), len(img.Data))
				// Show first few bytes to identify format
				if len(img.Data) > 16 {
					fmt.Printf("  Header bytes: %X\n", img.Data[:16])
				}
			}
		}
	}

	// Check the PPTX output
	fmt.Println("\n=== PPTX Output ===")
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	// Check rels for slide71
	for _, f := range r.File {
		if f.Name == "ppt/slides/_rels/slide71.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			content := string(data)
			rc.Close()
			fmt.Printf("slide71 rels:\n%s\n", content)
		}
	}

	// Check what image files exist
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			// Check if it's referenced by slide71
			name := f.Name[len("ppt/media/"):]
			if strings.Contains(name, "image165") || strings.Contains(name, "image166") ||
				strings.Contains(name, "img165") || strings.Contains(name, "img166") {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				fmt.Printf("\n%s: %d bytes\n", f.Name, len(data))
				if len(data) > 16 {
					fmt.Printf("  Header: %X\n", data[:16])
				}
			}
		}
	}
}
