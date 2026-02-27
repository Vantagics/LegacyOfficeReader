package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Parse the PPT
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPT: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	images := p.GetImages()
	masters := p.GetMasters()
	w, h := p.GetSlideSize()

	fmt.Printf("PPT: %d slides, %d images, %d masters, size=%dx%d\n", len(slides), len(images), len(masters), w, h)

	// Check PPTX
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	slideCount := 0
	layoutCount := 0
	imageCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			slideCount++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			layoutCount++
		}
		if strings.HasPrefix(f.Name, "ppt/media/") {
			imageCount++
		}
	}
	fmt.Printf("PPTX: %d slides, %d layouts, %d images\n", slideCount, layoutCount, imageCount)

	// Check each slide for issues
	fmt.Println("\n=== Slide-by-slide comparison ===")
	for i, s := range slides {
		shapes := s.GetShapes()
		bg := s.GetBackground()
		masterRef := s.GetMasterRef()

		// Count text shapes and image shapes
		textShapes := 0
		imageShapes := 0
		totalText := 0
		for _, sh := range shapes {
			if sh.IsImage {
				imageShapes++
			}
			if sh.IsText || len(sh.Paragraphs) > 0 {
				textShapes++
			}
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					totalText += len(r.Text)
				}
			}
		}

		// Check for potential issues
		issues := []string{}
		
		// Check for shapes with zero dimensions
		for j, sh := range shapes {
			if sh.Width == 0 && sh.Height == 0 && !isConnector(sh.ShapeType) {
				issues = append(issues, fmt.Sprintf("shape %d has zero size", j))
			}
			if sh.Left < -1000000 || sh.Top < -1000000 {
				issues = append(issues, fmt.Sprintf("shape %d has very negative position (%d,%d)", j, sh.Left, sh.Top))
			}
			// Check for missing font
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					if r.FontSize == 0 && r.Text != "" {
						issues = append(issues, fmt.Sprintf("shape %d has text with fontSize=0", j))
						break
					}
				}
			}
		}

		if len(issues) > 0 || i < 3 || i == len(slides)-1 {
			fmt.Printf("Slide %d: masterRef=%d, shapes=%d (text=%d, img=%d), bg=%v, textLen=%d\n",
				i+1, masterRef, len(shapes), textShapes, imageShapes, bg.HasBackground, totalText)
			for _, issue := range issues {
				fmt.Printf("  ISSUE: %s\n", issue)
			}
		}
	}

	// Check PPTX slide XML for validity
	fmt.Println("\n=== PPTX XML validation ===")
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			fmt.Printf("  %s: ERROR opening: %v\n", f.Name, err)
			continue
		}
		data, _ := io.ReadAll(rc)
		rc.Close()

		// Try to parse as XML
		decoder := xml.NewDecoder(strings.NewReader(string(data)))
		for {
			_, err := decoder.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				fmt.Printf("  %s: XML ERROR: %v\n", f.Name, err)
				break
			}
		}
	}
	fmt.Println("XML validation complete")

	// Check layout XML
	fmt.Println("\n=== Layout analysis ===")
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			continue
		}
		data, _ := io.ReadAll(rc)
		rc.Close()

		content := string(data)
		spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:pic>") + strings.Count(content, "<p:cxnSp>")
		hasBg := strings.Contains(content, "<p:bg>")
		hasBlipBg := strings.Contains(content, "blipFill") && strings.Contains(content, "<p:bg>")
		fmt.Printf("  %s: shapes=%d, hasBg=%v, hasBlipBg=%v\n", f.Name, spCount, hasBg, hasBlipBg)
	}

	// Check for shapes with text that might be cut off
	fmt.Println("\n=== Text overflow check (first 10 slides) ===")
	for i := 0; i < 10 && i < len(slides); i++ {
		s := slides[i]
		for j, sh := range s.GetShapes() {
			if !sh.IsText && len(sh.Paragraphs) == 0 {
				continue
			}
			totalChars := 0
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					totalChars += len(r.Text)
				}
			}
			if totalChars > 200 {
				fmt.Printf("  Slide %d, Shape %d: %d chars, size=(%d,%d)\n", i+1, j, totalChars, sh.Width, sh.Height)
			}
		}
	}
}

func isConnector(shapeType uint16) bool {
	switch shapeType {
	case 20, 32, 33, 34, 35, 36, 37, 38, 39, 40:
		return true
	}
	return false
}
