package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		panic(err)
	}
	defer f.Close()
	p, err := ppt.OpenReader(f)
	if err != nil {
		panic(err)
	}

	slides := p.GetSlides()
	if len(slides) < 4 {
		return
	}
	s := slides[3] // slide 4
	shapes := s.GetShapes()
	bg := s.GetBackground()
	scheme := s.GetColorScheme()

	fmt.Printf("Slide 4: %d shapes\n", len(shapes))
	fmt.Printf("Background: has=%v fill=%q imgIdx=%d\n", bg.HasBackground, bg.FillColor, bg.ImageIdx)
	fmt.Printf("ColorScheme: %v\n", scheme)

	// Print ALL shapes, especially non-text ones that might be background elements
	for si, sh := range shapes {
		text := ""
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				text += run.Text
			}
		}
		if len([]rune(text)) > 30 {
			text = string([]rune(text)[:30]) + "..."
		}

		kind := "SHAPE"
		if sh.IsImage {
			kind = fmt.Sprintf("IMAGE(idx=%d)", sh.ImageIdx)
		} else if sh.IsText {
			kind = "TEXT"
		}

		// Only print first 30 shapes and any with fill colors
		if si < 30 || sh.FillColor != "" || sh.IsImage {
			fmt.Printf("[%d] type=%d %s pos=(%d,%d) sz=(%d,%d) fill=%q noFill=%v fillRaw=0x%08X opacity=%d line=%q text=%q\n",
				si, sh.ShapeType, kind, sh.Left, sh.Top, sh.Width, sh.Height,
				sh.FillColor, sh.NoFill, sh.FillColorRaw, sh.FillOpacity,
				sh.LineColor, text)
		}
	}

	// Also check slide 5
	fmt.Println("\n=== Slide 5 shapes ===")
	s5 := slides[4]
	shapes5 := s5.GetShapes()
	for si, sh := range shapes5 {
		text := ""
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				text += run.Text
			}
		}
		if len([]rune(text)) > 30 {
			text = string([]rune(text)[:30]) + "..."
		}
		kind := "SHAPE"
		if sh.IsImage {
			kind = fmt.Sprintf("IMAGE(idx=%d)", sh.ImageIdx)
		} else if sh.IsText {
			kind = "TEXT"
		}
		fmt.Printf("[%d] type=%d %s pos=(%d,%d) sz=(%d,%d) fill=%q noFill=%v fillRaw=0x%08X opacity=%d text=%q\n",
			si, sh.ShapeType, kind, sh.Left, sh.Top, sh.Width, sh.Height,
			sh.FillColor, sh.NoFill, sh.FillColorRaw, sh.FillOpacity, text)
	}
}
