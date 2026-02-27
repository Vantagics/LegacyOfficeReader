package main

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Analyze PPT source
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
	images := p.GetImages()
	masters := p.GetMasters()
	sw, sh := p.GetSlideSize()

	fmt.Printf("=== PPT Source Analysis ===\n")
	fmt.Printf("Slides: %d, Images: %d, Masters: %d\n", len(slides), len(images), len(masters))
	fmt.Printf("Slide size: %d x %d EMU (%.1f x %.1f inches)\n", sw, sh, float64(sw)/914400, float64(sh)/914400)

	// Analyze masters
	for ref, m := range masters {
		fmt.Printf("\nMaster ref=%d: bg=%v shapes=%d colorScheme=%v\n",
			ref, m.Background.HasBackground, len(m.Shapes), m.ColorScheme)
		for i, sh := range m.Shapes {
			desc := "shape"
			if sh.IsImage {
				desc = fmt.Sprintf("IMAGE(idx=%d)", sh.ImageIdx)
			} else if sh.IsText && len(sh.Paragraphs) > 0 {
				text := ""
				for _, p := range sh.Paragraphs {
					for _, r := range p.Runs {
						text += r.Text
					}
				}
				if len(text) > 40 {
					text = text[:40]
				}
				desc = fmt.Sprintf("TEXT(%q)", text)
			}
			pct := float64(sh.Width) * float64(sh.Height) / (float64(sw) * float64(sh.Height))
			fmt.Printf("  shape[%d]: type=%d pos=(%d,%d) size=(%d,%d) fill=%q line=%q noFill=%v %s areaPct=%.1f%%\n",
				i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
				sh.FillColor, sh.LineColor, sh.NoFill, desc, pct*100)
		}
	}

	// Analyze first few slides in detail
	for i := 0; i < len(slides) && i < 5; i++ {
		s := slides[i]
		shapes := s.GetShapes()
		bg := s.GetBackground()
		fmt.Printf("\n--- Slide %d (layout=%d, masterRef=%d) ---\n", i+1, s.GetLayoutType(), s.GetMasterRef())
		fmt.Printf("  Background: has=%v fill=%q imgIdx=%d\n", bg.HasBackground, bg.FillColor, bg.ImageIdx)
		fmt.Printf("  Shapes: %d\n", len(shapes))
		for j, sh := range shapes {
			desc := ""
			if sh.IsImage {
				desc = fmt.Sprintf("IMAGE(idx=%d)", sh.ImageIdx)
			} else if len(sh.Paragraphs) > 0 {
				text := ""
				for _, p := range sh.Paragraphs {
					for _, r := range p.Runs {
						text += r.Text
					}
				}
				if len(text) > 60 {
					text = text[:60]
				}
				desc = fmt.Sprintf("TEXT(%q)", text)
			}
			fmt.Printf("  [%d] type=%d pos=(%d,%d) size=(%d,%d) fill=%q line=%q noFill=%v noLine=%v %s\n",
				j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
				sh.FillColor, sh.LineColor, sh.NoFill, sh.NoLine, desc)
			// Show text details
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					if run.Text == "" {
						continue
					}
					text := run.Text
					if len(text) > 40 {
						text = text[:40]
					}
					fmt.Printf("    p%d.r%d: font=%q sz=%d color=%q bold=%v %q\n",
						pi, ri, run.FontName, run.FontSize, run.Color, run.Bold, text)
				}
			}
		}
	}

	// Check PPTX output
	fmt.Printf("\n=== PPTX Output Analysis ===\n")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Printf("ERROR opening PPTX: %v\n", err)
		return
	}
	defer zr.Close()

	slideCount := 0
	layoutCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			slideCount++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			layoutCount++
		}
	}
	fmt.Printf("PPTX slides: %d, layouts: %d\n", slideCount, layoutCount)

	// Check slide 1 XML for issues
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			buf := make([]byte, 8000)
			n, _ := rc.Read(buf)
			rc.Close()
			content := string(buf[:n])
			fmt.Printf("\nSlide 1 XML (first 8000 chars):\n%s\n", content)
		}
		if f.Name == "ppt/slideLayouts/slideLayout5.xml" {
			rc, _ := f.Open()
			buf := make([]byte, 8000)
			n, _ := rc.Read(buf)
			rc.Close()
			content := string(buf[:n])
			fmt.Printf("\nLayout 5 XML (first 8000 chars):\n%s\n", content)
		}
	}

	// Analyze scheme color resolution
	fmt.Printf("\n=== Color Scheme Analysis ===\n")
	for i := 0; i < len(slides) && i < 10; i++ {
		s := slides[i]
		shapes := s.GetShapes()
		cs := s.GetColorScheme()
		noColor := 0
		totalRuns := 0
		for _, sh := range shapes {
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					if r.Text == "" {
						continue
					}
					totalRuns++
					if r.Color == "" {
						noColor++
					}
				}
			}
		}
		fmt.Printf("  Slide %d: runs=%d noColor=%d scheme=%v\n", i+1, totalRuns, noColor, cs)
	}
}
