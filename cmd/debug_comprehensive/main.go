package main

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Parse the original PPT
	pres, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPT: %v\n", err)
		os.Exit(1)
	}

	slides := pres.GetSlides()
	images := pres.GetImages()
	masters := pres.GetMasters()
	slideW, slideH := pres.GetSlideSize()

	fmt.Printf("=== PPT Summary ===\n")
	fmt.Printf("Slides: %d, Images: %d, Masters: %d\n", len(slides), len(images), len(masters))
	fmt.Printf("Slide size: %d x %d EMU\n", slideW, slideH)

	// Analyze each layout
	fmt.Printf("\n=== Layout Analysis ===\n")
	masterRefToSlides := make(map[uint32][]int)
	for i, s := range slides {
		ref := s.GetMasterRef()
		masterRefToSlides[ref] = append(masterRefToSlides[ref], i+1)
	}

	for ref, slideNums := range masterRefToSlides {
		m, ok := masters[ref]
		fmt.Printf("\nLayout ref=%d, slides=%d (first few: ", ref, len(slideNums))
		max := 5
		if len(slideNums) < max {
			max = len(slideNums)
		}
		for i := 0; i < max; i++ {
			if i > 0 {
				fmt.Print(", ")
			}
			fmt.Printf("%d", slideNums[i])
		}
		if len(slideNums) > 5 {
			fmt.Print("...")
		}
		fmt.Println(")")

		if ok {
			fmt.Printf("  Background: has=%v, fill=%s, imgIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
			fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
			fmt.Printf("  Shapes: %d\n", len(m.Shapes))
			for si, sh := range m.Shapes {
				desc := fmt.Sprintf("type=%d pos=(%d,%d) size=(%d,%d)", sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height)
				if sh.IsImage {
					desc += fmt.Sprintf(" IMG[%d]", sh.ImageIdx)
				}
				if sh.IsText {
					desc += " TEXT"
					for _, p := range sh.Paragraphs {
						for _, r := range p.Runs {
							t := strings.TrimSpace(r.Text)
							if t != "" && len(t) > 30 {
								t = t[:30] + "..."
							}
							if t != "" {
								desc += fmt.Sprintf(" \"%s\"", t)
							}
						}
					}
				}
				if sh.FillColor != "" {
					desc += fmt.Sprintf(" fill=%s", sh.FillColor)
				}
				if sh.LineColor != "" {
					desc += fmt.Sprintf(" line=%s", sh.LineColor)
				}
				fmt.Printf("  Shape[%d]: %s\n", si, desc)
			}
			// Default text styles
			for lvl, ts := range m.DefaultTextStyles {
				if ts.FontSize > 0 || ts.FontName != "" {
					fmt.Printf("  DefaultTextStyle[%d]: sz=%d font=%s bold=%v color=%s\n", lvl, ts.FontSize, ts.FontName, ts.Bold, ts.Color)
				}
			}
		}
	}

	// Analyze specific slides for issues
	fmt.Printf("\n=== Slide-by-Slide Analysis (first 10 + problematic) ===\n")
	checkSlides := []int{1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 41, 63, 71}
	for _, sn := range checkSlides {
		if sn > len(slides) {
			continue
		}
		s := slides[sn-1]
		shapes := s.GetShapes()
		bg := s.GetBackground()
		ref := s.GetMasterRef()
		fmt.Printf("\nSlide %d: shapes=%d, bg=(has=%v fill=%s img=%d), masterRef=%d\n",
			sn, len(shapes), bg.HasBackground, bg.FillColor, bg.ImageIdx, ref)

		// Count shape types
		textCount := 0
		imgCount := 0
		connCount := 0
		otherCount := 0
		for _, sh := range shapes {
			if sh.IsImage {
				imgCount++
			} else if isConnector(sh.ShapeType) {
				connCount++
			} else if sh.IsText || len(sh.Paragraphs) > 0 {
				textCount++
			} else {
				otherCount++
			}
		}
		fmt.Printf("  text=%d, img=%d, conn=%d, other=%d\n", textCount, imgCount, connCount, otherCount)

		// Show first few text shapes with their properties
		shown := 0
		for _, sh := range shapes {
			if shown >= 5 {
				break
			}
			if len(sh.Paragraphs) == 0 {
				continue
			}
			shown++
			text := ""
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					text += r.Text
				}
			}
			if len(text) > 60 {
				text = text[:60] + "..."
			}
			fmt.Printf("  TextShape: type=%d pos=(%d,%d) sz=(%d,%d) fill=%s noFill=%v\n",
				sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.NoFill)
			// Show first run properties
			if len(sh.Paragraphs) > 0 && len(sh.Paragraphs[0].Runs) > 0 {
				r := sh.Paragraphs[0].Runs[0]
				fmt.Printf("    Run: font=%s sz=%d bold=%v color=%s colorRaw=0x%08X\n",
					r.FontName, r.FontSize, r.Bold, r.Color, r.ColorRaw)
			}
			fmt.Printf("    Text: %s\n", text)
		}
	}

	// Check the generated PPTX
	fmt.Printf("\n=== PPTX Verification ===\n")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPTX: %v\n", err)
		return
	}
	defer zr.Close()

	slideCount := 0
	layoutCount := 0
	mediaCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideCount++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			layoutCount++
		}
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaCount++
		}
	}
	fmt.Printf("PPTX: slides=%d, layouts=%d, media=%d\n", slideCount, layoutCount, mediaCount)

	// Check a few slide XMLs for common issues
	for _, sn := range []int{1, 2, 5, 10} {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", sn)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, err := f.Open()
				if err != nil {
					continue
				}
				buf := make([]byte, 2000)
				n, _ := rc.Read(buf)
				rc.Close()
				content := string(buf[:n])
				// Check for showMasterSp
				hasMasterSp := strings.Contains(content, `showMasterSp="1"`)
				// Check for background
				hasBg := strings.Contains(content, `<p:bg>`)
				fmt.Printf("Slide %d: showMasterSp=%v, hasBg=%v, size=%d bytes\n", sn, hasMasterSp, hasBg, f.UncompressedSize64)
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
