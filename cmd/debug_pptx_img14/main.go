package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"image/png"
	"io"
	"os"
)

func main() {
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
			fmt.Printf("image14.png: %d bytes\n", len(data))

			img, err := png.Decode(bytes.NewReader(data))
			if err != nil {
				fmt.Printf("PNG decode error: %v\n", err)
			} else {
				bounds := img.Bounds()
				fmt.Printf("PNG dimensions: %dx%d\n", bounds.Dx(), bounds.Dy())
			}

			// Compare with source
			srcData, _ := os.ReadFile("testfie/watermark_check.png")
			if bytes.Equal(data, srcData) {
				fmt.Println("MATCH: image14.png matches source watermark")
			} else {
				fmt.Printf("MISMATCH: image14.png (%d bytes) vs source (%d bytes)\n", len(data), len(srcData))
			}
		}
	}
}
