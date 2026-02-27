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
	// Check the blip data for image index 13 (the watermark)
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

	images := p.GetImages()
	fmt.Printf("Total images: %d\n", len(images))

	if len(images) > 13 {
		img := images[13]
		fmt.Printf("Image 13: format=%s, size=%d bytes\n", img.Format, len(img.Data))
		if len(img.Data) > 8 {
			fmt.Printf("  First 8 bytes: %02X %02X %02X %02X %02X %02X %02X %02X\n",
				img.Data[0], img.Data[1], img.Data[2], img.Data[3],
				img.Data[4], img.Data[5], img.Data[6], img.Data[7])
		}
	}

	// Also check what's in the PPTX
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, zf := range r.File {
		if zf.Name == "ppt/media/image14.png" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\nPPTX image14.png: %d bytes\n", len(data))
			if len(data) > 8 {
				fmt.Printf("  First 8 bytes: %02X %02X %02X %02X %02X %02X %02X %02X\n",
					data[0], data[1], data[2], data[3],
					data[4], data[5], data[6], data[7])
				// Check if it's actually PNG
				if data[0] == 0x89 && data[1] == 0x50 {
					fmt.Println("  Format: PNG")
				} else if data[0] == 0x01 && data[1] == 0x00 {
					fmt.Println("  Format: EMF (not PNG!)")
				} else if data[0] == 0xD7 && data[1] == 0xC6 {
					fmt.Println("  Format: WMF (not PNG!)")
				}
			}
		}
	}

	// Check the watermark position in the PPTX XML
	for _, zf := range r.File {
		if zf.Name != "ppt/slides/slide8.xml" {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Find the pic element with rImg14
		idx := strings.Index(content, `rImg14`)
		if idx >= 0 {
			// Show surrounding context
			start := idx - 200
			if start < 0 {
				start = 0
			}
			end := idx + 200
			if end > len(content) {
				end = len(content)
			}
			fmt.Printf("\nContext around rImg14:\n%s\n", content[start:end])
		}
	}
}
