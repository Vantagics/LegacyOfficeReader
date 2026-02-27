package main

import (
	"bytes"
	"fmt"
	"image/png"
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
	// Image index 13 is the tiger watermark
	img := images[13]
	fmt.Printf("Image 13: format=%d size=%d bytes\n", img.Format, len(img.Data))

	// Try to decode as PNG to get dimensions
	r := bytes.NewReader(img.Data)
	pngImg, err := png.Decode(r)
	if err != nil {
		fmt.Printf("Not a valid PNG: %v\n", err)
		// Check first bytes
		if len(img.Data) > 8 {
			fmt.Printf("First 8 bytes: %x\n", img.Data[:8])
		}
	} else {
		bounds := pngImg.Bounds()
		fmt.Printf("PNG dimensions: %dx%d\n", bounds.Dx(), bounds.Dy())
	}

	// Also save it to check visually
	os.WriteFile("testfie/watermark_check.png", img.Data, 0644)
	fmt.Println("Saved to testfie/watermark_check.png")
}
