package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "DOC: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	// Show all paragraphs with drawn images
	for i, p := range fc.Paragraphs {
		if len(p.DrawnImages) > 0 || p.TextBoxText != "" {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			fmt.Printf("P[%d]: drawn=%v textbox=%q heading=%d text=%q\n", i, p.DrawnImages, p.TextBoxText, p.HeadingLevel, text)
		}
	}

	// Also show paragraphs with inline images
	for i, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				fmt.Printf("P[%d]: inline imageRef=%d text=%q\n", i, r.ImageRef, r.Text)
			}
		}
	}
}
