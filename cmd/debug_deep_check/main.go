package main

import (
	"archive/zip"
	"fmt"
	"os"
	"regexp"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Parse the original PPT
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "PPT open error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()
	fmt.Printf("PPT: %d slides, %d masters\n", len(slides), len(masters))

	// Check color scheme
	for id, m := range masters {
		fmt.Printf("Master %d: scheme=%v\n", id, m.ColorScheme)
		fmt.Printf("  DefaultTextStyles:\n")
		for i, s := range m.DefaultTextStyles {
			if s.FontSize > 0 || s.Color != "" {
				fmt.Printf("    Level %d: size=%d font=%q color=%s bold=%v\n",
					i, s.FontSize, s.FontName, s.Color, s.Bold)
			}
		}
		fmt.Printf("  Shapes: %d\n", len(m.Shapes))
		for i, sh := range m.Shapes {
			if sh.FillColor != "" || sh.LineColor != "" || sh.IsImage {
				fmt.Printf("    Shape %d: type=%d fill=%s line=%s img=%v imgIdx=%d pos=(%d,%d) size=(%d,%d)\n",
					i, sh.ShapeType, sh.FillColor, sh.LineColor, sh.IsImage, sh.ImageIdx,
					sh.Left, sh.Top, sh.Width, sh.Height)
			}
		}
	}

	// Check first few slides for text content and colors
	fmt.Printf("\n=== Slide Content Check ===\n")
	checkSlides := []int{0, 1, 2, 3, 9, 15, 20, 27} // section dividers and content slides
	for _, idx := range checkSlides {
		if idx >= len(slides) {
			continue
		}
		s := slides[idx]
		shapes := s.GetShapes()
		bg := s.GetBackground()
		fmt.Printf("\nSlide %d: %d shapes, bg=%v bgColor=%s bgImg=%d masterRef=%d\n",
			idx+1, len(shapes), bg.HasBackground, bg.FillColor, bg.ImageIdx, s.GetMasterRef())

		for si, sh := range shapes {
			if len(sh.Paragraphs) > 0 {
				fmt.Printf("  Shape %d (type=%d, fill=%s, %dx%d):\n",
					si, sh.ShapeType, sh.FillColor, sh.Width, sh.Height)
				for pi, para := range sh.Paragraphs {
					if pi > 3 {
						fmt.Printf("    ... (%d more paragraphs)\n", len(sh.Paragraphs)-pi)
						break
					}
					for _, run := range para.Runs {
						text := run.Text
						if len(text) > 60 {
							text = text[:60] + "..."
						}
						fmt.Printf("    [%s] sz=%d font=%q color=%s bold=%v\n",
							text, run.FontSize, run.FontName, run.Color, run.Bold)
					}
				}
			}
		}
	}

	// Now check the PPTX output for matching content
	fmt.Printf("\n=== PPTX Output Verification ===\n")
	zf, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "PPTX open error: %v\n", err)
		os.Exit(1)
	}
	defer zf.Close()

	szRe := regexp.MustCompile(`sz="(\d+)"`)
	colorRe := regexp.MustCompile(`<a:srgbClr val="([A-F0-9]{6})"`)

	for _, idx := range checkSlides {
		slideName := fmt.Sprintf("ppt/slides/slide%d.xml", idx+1)
		for _, file := range zf.File {
			if file.Name != slideName {
				continue
			}
			rc, _ := file.Open()
			buf := make([]byte, file.UncompressedSize64)
			n, _ := rc.Read(buf)
			rc.Close()
			content := string(buf[:n])

			// Extract font sizes
			sizes := szRe.FindAllStringSubmatch(content, -1)
			sizeMap := make(map[string]int)
			for _, m := range sizes {
				sizeMap[m[1]]++
			}

			// Extract colors
			colors := colorRe.FindAllStringSubmatch(content, -1)
			colorMap := make(map[string]int)
			for _, m := range colors {
				colorMap[m[1]]++
			}

			fmt.Printf("\nSlide %d PPTX:\n", idx+1)
			fmt.Printf("  Font sizes: ")
			for sz, cnt := range sizeMap {
				fmt.Printf("%s(%d) ", sz, cnt)
			}
			fmt.Printf("\n  Colors: ")
			for c, cnt := range colorMap {
				fmt.Printf("%s(%d) ", c, cnt)
			}
			fmt.Println()

			// Check for suspicious patterns
			if strings.Contains(content, `val="000008"`) || strings.Contains(content, `val="0000FE"`) {
				fmt.Printf("  WARNING: Possible unresolved scheme color reference!\n")
			}
		}
	}
}
