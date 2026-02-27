package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	document, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	fc := document.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Find all paragraphs with inline images (runs with \x01 or ImageRef >= 0)
	fmt.Println("=== Paragraphs with inline images ===")
	for i, p := range fc.Paragraphs {
		for j, r := range p.Runs {
			if strings.Contains(r.Text, "\x01") || r.ImageRef >= 0 {
				text := r.Text
				if len(text) > 60 {
					text = text[:60] + "..."
				}
				fmt.Printf("P[%d] Run[%d]: ImageRef=%d hasPicLoc=%v picLoc=%d text=%q\n",
					i, j, r.ImageRef, r.Props.HasPicLocation, r.Props.PicLocation, text)
			}
		}
	}

	// Find all paragraphs with drawn images
	fmt.Println("\n=== Paragraphs with drawn images ===")
	for i, p := range fc.Paragraphs {
		if len(p.DrawnImages) > 0 {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 80 {
				text = text[:80] + "..."
			}
			fmt.Printf("P[%d] DrawnImages=%v: %q\n", i, p.DrawnImages, text)
		}
	}

	// Find all paragraphs with text boxes
	fmt.Println("\n=== Paragraphs with text boxes ===")
	for i, p := range fc.Paragraphs {
		if p.TextBoxText != "" {
			fmt.Printf("P[%d] TextBox=%q\n", i, p.TextBoxText)
		}
	}

	// Show image details
	images := document.GetImages()
	fmt.Println("\n=== Image details ===")
	for i, img := range images {
		ext := ".bin"
		switch img.Format {
		case 0:
			ext = ".emf"
		case 4:
			ext = ".png"
		}
		fmt.Printf("BSE[%d]: format=%s size=%d bytes\n", i, ext, len(img.Data))
	}

	// Count total \x01 characters across all runs
	total01 := 0
	for _, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			total01 += strings.Count(r.Text, "\x01")
		}
	}
	fmt.Printf("\nTotal \\x01 characters: %d\n", total01)
}
