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

	// Check shape counts and identify shapes that might be filtered
	for i, s := range slides {
		shapes := s.GetShapes()
		if len(shapes) < 10 {
			continue // skip simple slides
		}

		textShapes := 0
		imageShapes := 0
		connectors := 0
		emptyText := 0
		otherShapes := 0
		zeroSize := 0

		for _, sh := range shapes {
			if sh.Width == 0 && sh.Height == 0 {
				zeroSize++
			}
			if sh.IsImage && sh.ImageIdx >= 0 {
				imageShapes++
			} else if isConnector(sh.ShapeType) {
				connectors++
			} else if sh.IsText {
				hasText := false
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						if strings.TrimSpace(run.Text) != "" {
							hasText = true
						}
					}
				}
				if hasText {
					textShapes++
				} else {
					emptyText++
				}
			} else {
				otherShapes++
			}
		}

		fmt.Printf("Slide %2d: total=%3d text=%2d emptyText=%2d img=%2d conn=%2d other=%2d zeroSize=%d\n",
			i+1, len(shapes), textShapes, emptyText, imageShapes, connectors, otherShapes, zeroSize)
	}

	// Detailed look at slide 4 shapes
	fmt.Println("\n--- Slide 4 Shape Details ---")
	shapes := slides[3].GetShapes()
	typeCounts := make(map[uint16]int)
	for _, sh := range shapes {
		typeCounts[sh.ShapeType]++
	}
	for t, c := range typeCounts {
		fmt.Printf("  ShapeType %d: %d shapes\n", t, c)
	}

	// Show first 20 shapes
	for j, sh := range shapes {
		if j >= 30 {
			break
		}
		text := ""
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				t := strings.TrimSpace(run.Text)
				if t != "" {
					text += t + " "
				}
			}
		}
		if len([]rune(text)) > 40 {
			text = string([]rune(text)[:40]) + "..."
		}
		fmt.Printf("  [%3d] type=%3d pos=(%d,%d) size=(%d,%d) text=%v img=%v imgIdx=%d fill=%q noFill=%v line=%q text=%q\n",
			j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
			sh.IsText, sh.IsImage, sh.ImageIdx, sh.FillColor, sh.NoFill, sh.LineColor, text)
	}

	// Check slide 9 (131 shapes)
	fmt.Println("\n--- Slide 9 Shape Details ---")
	shapes9 := slides[8].GetShapes()
	typeCounts9 := make(map[uint16]int)
	for _, sh := range shapes9 {
		typeCounts9[sh.ShapeType]++
	}
	for t, c := range typeCounts9 {
		fmt.Printf("  ShapeType %d: %d shapes\n", t, c)
	}
}

func isConnector(shapeType uint16) bool {
	switch shapeType {
	case 20, 32, 33, 34, 35, 36, 37, 38, 39, 40:
		return true
	}
	return false
}
