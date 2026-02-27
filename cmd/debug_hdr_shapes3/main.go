package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		panic(err)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	fmt.Printf("Headers: %d\n", len(fc.Headers))
	fmt.Printf("Footers: %d\n", len(fc.Footers))
	fmt.Printf("HeadersRaw: %d\n", len(fc.HeadersRaw))
	fmt.Printf("FootersRaw: %d\n", len(fc.FootersRaw))

	for i, h := range fc.Headers {
		fmt.Printf("Header[%d]: %q\n", i, h)
	}
	for i, h := range fc.HeadersRaw {
		fmt.Printf("HeaderRaw[%d]: %q\n", i, h)
	}
	for i, f := range fc.Footers {
		fmt.Printf("Footer[%d]: %q\n", i, f)
	}
	for i, f := range fc.FootersRaw {
		fmt.Printf("FooterRaw[%d]: %q\n", i, f)
	}

	// Check images
	images := d.GetImages()
	fmt.Printf("\nImages: %d\n", len(images))
	for i, img := range images {
		fmt.Printf("  [%d] format=%d size=%d\n", i, img.Format, len(img.Data))
	}
}
