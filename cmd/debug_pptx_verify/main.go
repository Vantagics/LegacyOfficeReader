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
	// Open original PPT
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPT: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	images := p.GetImages()
	sw, sh := p.GetSlideSize()

	fmt.Printf("PPT: %d slides, %d images, size=%dx%d\n", len(slides), len(images), sw, sh)

	// Per-slide shape count comparison
	fmt.Printf("\n=== Per-Slide Shape Counts ===\n")
	for i, s := range slides {
		shapes := s.GetShapes()
		bg := s.GetBackground()
		textCount := 0
		imgCount := 0
		otherCount := 0
		for _, sh := range shapes {
			if sh.IsImage {
				imgCount++
			} else if sh.IsText || len(sh.Paragraphs) > 0 {
				textCount++
			} else {
				otherCount++
			}
		}
		bgStr := "-"
		if bg.HasBackground {
			if bg.ImageIdx >= 0 {
				bgStr = fmt.Sprintf("img%d", bg.ImageIdx)
			} else {
				bgStr = bg.FillColor
			}
		}
		fmt.Printf("S%02d: total=%d text=%d img=%d other=%d bg=%s master=%d\n",
			i+1, len(shapes), textCount, imgCount, otherCount, bgStr, s.GetMasterRef())
	}

	// Verify PPTX
	fmt.Printf("\n=== PPTX Verification ===\n")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	pptxSlideCount := 0
	pptxLayoutCount := 0
	pptxImageCount := 0
	xmlErrors := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			pptxSlideCount++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			pptxLayoutCount++
		}
		if strings.HasPrefix(f.Name, "ppt/media/") {
			pptxImageCount++
		}
	}
	fmt.Printf("PPTX: %d slides, %d layouts, %d images\n", pptxSlideCount, pptxLayoutCount, pptxImageCount)

	// Validate all XML files
	for _, f := range zr.File {
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			fmt.Printf("ERROR opening %s: %v\n", f.Name, err)
			xmlErrors++
			continue
		}
		d := xml.NewDecoder(rc)
		for {
			_, err := d.Token()
			if err != nil {
				if err == io.EOF {
					break
				}
				fmt.Printf("XML ERROR in %s: %v\n", f.Name, err)
				xmlErrors++
				break
			}
		}
		rc.Close()
	}
	if xmlErrors == 0 {
		fmt.Println("All XML files valid")
	} else {
		fmt.Printf("%d XML errors found\n", xmlErrors)
	}

	// Count shapes per PPTX slide by parsing XML
	fmt.Printf("\n=== PPTX Slide Shape Counts ===\n")
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			continue
		}
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
		picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
		cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
		total := spCount + picCount + cxnCount

		// Extract slide number from filename
		slideNum := strings.TrimPrefix(f.Name, "ppt/slides/slide")
		slideNum = strings.TrimSuffix(slideNum, ".xml")

		fmt.Printf("slide%s: sp=%d pic=%d cxn=%d total=%d\n", slideNum, spCount, picCount, cxnCount, total)
	}

	// Check specific issues
	fmt.Printf("\n=== Issue Checks ===\n")

	// Check for shapes with out-of-range image indices
	outOfRange := 0
	for i, s := range slides {
		for j, sh := range s.GetShapes() {
			if sh.IsImage && (sh.ImageIdx < 0 || sh.ImageIdx >= len(images)) {
				fmt.Printf("Slide %d shape %d: imgIdx=%d out of range (max=%d)\n", i+1, j, sh.ImageIdx, len(images)-1)
				outOfRange++
			}
		}
	}
	if outOfRange == 0 {
		fmt.Println("No out-of-range image indices")
	}

	// Check for text runs with no font size after resolution
	noFontSize := 0
	for i, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.FontSize == 0 && strings.TrimSpace(run.Text) != "" {
						noFontSize++
						if noFontSize <= 5 {
							fmt.Printf("Slide %d: run with no fontSize: %q\n", i+1, truncStr(run.Text, 30))
						}
					}
				}
			}
		}
	}
	fmt.Printf("Runs with no font size: %d\n", noFontSize)

	// Check dark fill shapes for text color
	darkFillWhiteText := 0
	darkFillBlackText := 0
	for i, s := range slides {
		for _, sh := range s.GetShapes() {
			if sh.FillColor == "" || sh.NoFill {
				continue
			}
			if !isDark(sh.FillColor) {
				continue
			}
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if strings.TrimSpace(run.Text) == "" {
						continue
					}
					if run.Color == "FFFFFF" || run.Color == "ffffff" {
						darkFillWhiteText++
					} else if run.Color == "" || run.Color == "000000" {
						darkFillBlackText++
						if darkFillBlackText <= 5 {
							fmt.Printf("Slide %d: dark fill=%s, text color=%q: %q\n",
								i+1, sh.FillColor, run.Color, truncStr(run.Text, 30))
						}
					}
				}
			}
		}
	}
	fmt.Printf("Dark fill shapes: white text=%d, black/empty text=%d\n", darkFillWhiteText, darkFillBlackText)

	// Check yellow banner font sizes
	fmt.Printf("\n=== Yellow Banner Font Sizes ===\n")
	for i, s := range slides {
		for _, sh := range s.GetShapes() {
			if sh.FillColor != "FFD966" {
				continue
			}
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if strings.TrimSpace(run.Text) != "" {
						fmt.Printf("Slide %d: yellow banner sz=%d text=%q\n",
							i+1, run.FontSize, truncStr(run.Text, 40))
						break
					}
				}
			}
		}
	}
}

func truncStr(s string, n int) string {
	r := []rune(s)
	if len(r) > n {
		return string(r[:n]) + "..."
	}
	return s
}

func isDark(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	r := hexVal(hex[0])*16 + hexVal(hex[1])
	g := hexVal(hex[2])*16 + hexVal(hex[3])
	b := hexVal(hex[4])*16 + hexVal(hex[5])
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	return lum < 128
}

func hexVal(c byte) int {
	switch {
	case c >= '0' && c <= '9':
		return int(c - '0')
	case c >= 'A' && c <= 'F':
		return int(c-'A') + 10
	case c >= 'a' && c <= 'f':
		return int(c-'a') + 10
	}
	return 0
}
