package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, _ := os.Open("testfie/test.doc")
	defer f.Close()
	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	// Show paragraphs with inline images and their context
	for i, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			if r.ImageRef >= 0 || strings.Contains(r.Text, "\x01") {
				// Get surrounding text
				allText := ""
				for _, r2 := range p.Runs {
					allText += r2.Text
				}
				if len(allText) > 80 {
					allText = allText[:80] + "..."
				}
				fmt.Printf("P[%d]: imageRef=%d picLoc=%d hasPicLoc=%v text=%q\n",
					i, r.ImageRef, r.Props.PicLocation, r.Props.HasPicLocation, allText)
			}
		}
	}

	// Show images info
	images := d.GetImages()
	fmt.Printf("\nImages: %d total\n", len(images))
	for i, img := range images {
		fmt.Printf("  BSE[%d]: format=%d size=%d\n", i, img.Format, len(img.Data))
	}
}
