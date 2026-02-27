package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	slides := p.GetSlides()
	if len(slides) < 39 {
		fmt.Println("Not enough slides")
		return
	}

	s := slides[38] // 0-indexed
	shapes := s.GetShapes()
	fmt.Printf("Slide 39: %d shapes\n", len(shapes))
	for si, sh := range shapes {
		if len(sh.Paragraphs) == 0 {
			continue
		}
		fmt.Printf("Shape %d (type=%d):\n", si, sh.ShapeType)
		for pi, para := range sh.Paragraphs {
			for ri, run := range para.Runs {
				text := run.Text
				// Show control characters
				escaped := strings.ReplaceAll(text, "\n", "\\n")
				escaped = strings.ReplaceAll(escaped, "\r", "\\r")
				escaped = strings.ReplaceAll(escaped, "\x0b", "\\v")
				escaped = strings.ReplaceAll(escaped, "\t", "\\t")
				fmt.Printf("  P%d R%d: [%s] (len=%d)\n", pi, ri, escaped, len(text))
			}
		}
	}
}
