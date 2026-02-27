package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	slides := p.GetSlides()

	// Check slide 9 (131 shapes, 88 "other" shapes)
	// These "other" shapes are likely decorative rectangles forming a diagram
	fmt.Println("=== Slide 9 'other' shapes (non-text, non-image, non-connector) ===")
	shapes9 := slides[8].GetShapes()
	otherCount := 0
	for j, sh := range shapes9 {
		if sh.IsImage || sh.IsText || isConnector(sh.ShapeType) {
			continue
		}
		otherCount++
		if otherCount <= 15 {
			fmt.Printf("  [%d] type=%d pos=(%d,%d) size=(%d,%d) fill=%q noFill=%v line=%q lineW=%d\n",
				j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
				sh.FillColor, sh.NoFill, sh.LineColor, sh.LineWidth)
		}
	}
	fmt.Printf("  Total 'other' shapes: %d\n", otherCount)

	// Check slide 8 (15 "other" shapes)
	fmt.Println("\n=== Slide 8 'other' shapes ===")
	shapes8 := slides[7].GetShapes()
	for j, sh := range shapes8 {
		if sh.IsImage || sh.IsText || isConnector(sh.ShapeType) {
			continue
		}
		fmt.Printf("  [%d] type=%d pos=(%d,%d) size=(%d,%d) fill=%q noFill=%v line=%q lineW=%d\n",
			j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
			sh.FillColor, sh.NoFill, sh.LineColor, sh.LineWidth)
	}

	// Check slide 2 (8 "other" shapes - likely the CONTENTS page decorations)
	fmt.Println("\n=== Slide 2 'other' shapes ===")
	shapes2 := slides[1].GetShapes()
	for j, sh := range shapes2 {
		if sh.IsImage || isConnector(sh.ShapeType) {
			continue
		}
		text := ""
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				text += run.Text
			}
		}
		fmt.Printf("  [%d] type=%d pos=(%d,%d) size=(%d,%d) fill=%q noFill=%v isText=%v text=%q\n",
			j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
			sh.FillColor, sh.NoFill, sh.IsText, text)
	}
}

func isConnector(shapeType uint16) bool {
	switch shapeType {
	case 20, 32, 33, 34, 35, 36, 37, 38, 39, 40:
		return true
	}
	return false
}
