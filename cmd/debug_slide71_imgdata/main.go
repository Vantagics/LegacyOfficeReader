package main

import (
	"encoding/binary"
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

	// Check image 164 (0-indexed, Shape[1] references ImageIdx=164)
	for _, idx := range []int{164, 165} {
		if idx < len(images) {
			img := images[idx]
			fmt.Printf("Image[%d]: format=%d (%s), %d bytes\n", idx, img.Format, img.Extension(), len(img.Data))
			if len(img.Data) > 0 {
				fmt.Printf("  Full hex dump (first 64 bytes): %X\n", img.Data[:min(64, len(img.Data))])
			}
			// For PNG, parse IHDR to get dimensions
			if img.Format == 4 && len(img.Data) > 24 { // PNG
				w := binary.BigEndian.Uint32(img.Data[16:20])
				h := binary.BigEndian.Uint32(img.Data[20:24])
				fmt.Printf("  PNG dimensions: %dx%d\n", w, h)
			}
		}
	}

	// Also check nearby images to see if there's a pattern
	fmt.Println("\n--- Images around index 164 ---")
	for idx := 160; idx < min(167, len(images)); idx++ {
		img := images[idx]
		fmt.Printf("Image[%d]: format=%d (%s), %d bytes", idx, img.Format, img.Extension(), len(img.Data))
		if img.Format == 4 && len(img.Data) > 24 { // PNG
			w := binary.BigEndian.Uint32(img.Data[16:20])
			h := binary.BigEndian.Uint32(img.Data[20:24])
			fmt.Printf(" (%dx%d)", w, h)
		}
		fmt.Println()
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
