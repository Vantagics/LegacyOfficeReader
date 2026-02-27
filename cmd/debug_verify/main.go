package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	fmt.Printf("Total paragraphs: %d\n", len(fc.Paragraphs))

	// Show paragraphs with drawn images
	fmt.Println("\n=== Paragraphs with DrawnImages ===")
	for i, p := range fc.Paragraphs {
		if len(p.DrawnImages) > 0 {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			fmt.Printf("  P%d: DrawnImages=%v text=%q\n", i, p.DrawnImages, truncate(text, 60))
		}
	}

	// Show paragraphs with inline images (0x01 in runs)
	fmt.Println("\n=== Paragraphs with inline images ===")
	for i, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			if strings.Contains(r.Text, "\x01") {
				fmt.Printf("  P%d: ImageRef=%d text=%q\n", i, r.ImageRef, truncate(r.Text, 60))
			}
		}
	}

	// Show first 20 paragraphs (title page area)
	fmt.Println("\n=== First 20 paragraphs ===")
	for i := 0; i < 20 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		extra := ""
		if len(p.DrawnImages) > 0 {
			extra += fmt.Sprintf(" [DRAWN:%v]", p.DrawnImages)
		}
		if p.HasPageBreak {
			extra += " [PAGEBREAK]"
		}
		if p.HeadingLevel > 0 {
			extra += fmt.Sprintf(" [H%d]", p.HeadingLevel)
		}
		fmt.Printf("  P%d: %q%s\n", i, truncate(text, 60), extra)
	}

	// Show images
	images := d.GetImages()
	fmt.Printf("\n=== Images: %d total ===\n", len(images))
	for i, img := range images {
		ext := "?"
		switch img.Format {
		case 0:
			ext = "EMF"
		case 4:
			ext = "PNG"
		case 3:
			ext = "JPEG"
		}
		fmt.Printf("  BSE[%d]: %s %d bytes\n", i, ext, len(img.Data))
	}
}

func truncate(s string, n int) string {
	runes := []rune(s)
	if len(runes) > n {
		return string(runes[:n]) + "..."
	}
	return s
}
