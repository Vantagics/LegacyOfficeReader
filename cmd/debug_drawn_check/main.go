package main

import (
	"fmt"
	"os"

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

	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))

	// Show paragraphs with drawn images or text box
	for i, p := range fc.Paragraphs {
		if len(p.DrawnImages) > 0 || p.TextBoxText != "" {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			fmt.Printf("P[%d]: drawn=%v textbox=%q text=%q\n", i, p.DrawnImages, p.TextBoxText, text)
		}
	}

	// Show paragraphs with inline images
	for i, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				text := r.Text
				if len(text) > 40 {
					text = text[:40] + "..."
				}
				fmt.Printf("P[%d]: inline imageRef=%d text=%q\n", i, r.ImageRef, text)
			}
		}
	}

	// Count total images
	totalDrawn := 0
	for _, p := range fc.Paragraphs {
		totalDrawn += len(p.DrawnImages)
	}
	fmt.Printf("\nTotal drawn image refs: %d\n", totalDrawn)
}
