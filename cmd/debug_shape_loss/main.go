package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	// Open PPTX
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check slide 9 (0-indexed: 8) in detail
	// PPT has 131 shapes, PPTX has 49
	slideIdx := 8
	s := slides[slideIdx]
	shapes := s.GetShapes()

	fmt.Printf("=== Slide %d: %d PPT shapes ===\n", slideIdx+1, len(shapes))

	// Count shapes by category
	imgShapes := 0
	connShapes := 0
	textShapes := 0
	geomShapes := 0 // non-text, non-image, non-connector shapes

	for _, sh := range shapes {
		if sh.IsImage {
			imgShapes++
		} else if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
			connShapes++
		} else if sh.IsText || len(sh.Paragraphs) > 0 {
			textShapes++
		} else {
			geomShapes++
		}
	}

	fmt.Printf("  Images: %d, Connectors: %d, Text: %d, Geometry-only: %d\n",
		imgShapes, connShapes, textShapes, geomShapes)

	// The geometry-only shapes are the ones being lost!
	// Let's see what they look like
	fmt.Printf("\n=== Geometry-only shapes (no text, no image, not connector) ===\n")
	for i, sh := range shapes {
		if sh.IsImage || sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
			continue
		}
		if sh.IsText || len(sh.Paragraphs) > 0 {
			continue
		}
		fmt.Printf("  [%3d] type=%3d fill=%s noFill=%v line=%s noLine=%v w=%d h=%d\n",
			i, sh.ShapeType, sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine, sh.Width, sh.Height)
	}

	// Now check PPTX slide 9
	fname := fmt.Sprintf("ppt/slides/slide%d.xml", slideIdx+1)
	content := readZipFile(zr, fname)
	spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
	picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
	cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
	fmt.Printf("\nPPTX slide %d: sp=%d pic=%d cxn=%d total=%d\n",
		slideIdx+1, spCount, picCount, cxnCount, spCount+picCount+cxnCount)

	// Now check slide 4 (0-indexed: 3)
	slideIdx = 3
	s = slides[slideIdx]
	shapes = s.GetShapes()
	fmt.Printf("\n=== Slide %d: %d PPT shapes ===\n", slideIdx+1, len(shapes))

	imgShapes = 0
	connShapes = 0
	textShapes = 0
	geomShapes = 0
	for _, sh := range shapes {
		if sh.IsImage {
			imgShapes++
		} else if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
			connShapes++
		} else if sh.IsText || len(sh.Paragraphs) > 0 {
			textShapes++
		} else {
			geomShapes++
		}
	}
	fmt.Printf("  Images: %d, Connectors: %d, Text: %d, Geometry-only: %d\n",
		imgShapes, connShapes, textShapes, geomShapes)

	// Check: are connectors with width=0 or height=0 being written?
	fmt.Printf("\n=== Connectors with zero dimension ===\n")
	zeroWConn := 0
	zeroHConn := 0
	for _, sh := range shapes {
		if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
			if sh.Width == 0 {
				zeroWConn++
			}
			if sh.Height == 0 {
				zeroHConn++
			}
		}
	}
	fmt.Printf("  Connectors with width=0: %d, height=0: %d\n", zeroWConn, zeroHConn)

	// Check slide 26 (0-indexed: 25) - PPT=175, PPTX=47
	slideIdx = 25
	s = slides[slideIdx]
	shapes = s.GetShapes()
	fmt.Printf("\n=== Slide %d: %d PPT shapes ===\n", slideIdx+1, len(shapes))

	imgShapes = 0
	connShapes = 0
	textShapes = 0
	geomShapes = 0
	for _, sh := range shapes {
		if sh.IsImage {
			imgShapes++
		} else if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
			connShapes++
		} else if sh.IsText || len(sh.Paragraphs) > 0 {
			textShapes++
		} else {
			geomShapes++
		}
	}
	fmt.Printf("  Images: %d, Connectors: %d, Text: %d, Geometry-only: %d\n",
		imgShapes, connShapes, textShapes, geomShapes)

	// The PPTX has 47 shapes but PPT has 175. That's 128 missing.
	// connectors=31, so 175-31=144 non-connector shapes, but PPTX has 47 total
	// Let's check: 47 = sp + pic + cxn. If cxn=0 in PPTX...
	content = readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", slideIdx+1))
	spCount = strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
	picCount = strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
	cxnCount = strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
	fmt.Printf("PPTX slide %d: sp=%d pic=%d cxn=%d total=%d\n",
		slideIdx+1, spCount, picCount, cxnCount, spCount+picCount+cxnCount)

	// Check: how many text shapes have multiple paragraphs that could be table cells?
	// Slide 26 is a table-like layout with many small text cells
	fmt.Printf("\n=== Slide 26 text shape sizes ===\n")
	smallShapes := 0
	for _, sh := range shapes {
		if sh.IsText || len(sh.Paragraphs) > 0 {
			if sh.Height < 300000 { // less than ~0.6cm
				smallShapes++
			}
		}
	}
	fmt.Printf("  Small text shapes (h<300000 EMU): %d\n", smallShapes)
}

func readZipFile(zr *zip.ReadCloser, name string) string {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			return string(data)
		}
	}
	return ""
}
