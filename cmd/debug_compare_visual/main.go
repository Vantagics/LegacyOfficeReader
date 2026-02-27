package main

import (
	"archive/zip"
	"fmt"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Parse the PPT
	presentation, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	slides := presentation.GetSlides()
	masters := presentation.GetMasters()
	images := presentation.GetImages()

	fmt.Printf("=== PPT Summary ===\n")
	fmt.Printf("Slides: %d, Images: %d, Masters: %d\n", len(slides), len(images), len(masters))

	slideW, slideH := presentation.GetSlideSize()
	fmt.Printf("Slide size: %d x %d EMU\n", slideW, slideH)

	// Analyze each slide
	for i, s := range slides {
		shapes := s.GetShapes()
		bg := s.GetBackground()
		masterRef := s.GetMasterRef()

		fmt.Printf("\n--- Slide %d (masterRef=%d, shapes=%d) ---\n", i+1, masterRef, len(shapes))
		if bg.HasBackground {
			fmt.Printf("  BG: color=%s imgIdx=%d\n", bg.FillColor, bg.ImageIdx)
		}

		for j, sh := range shapes {
			desc := ""
			if sh.IsImage {
				desc = fmt.Sprintf("IMAGE(idx=%d)", sh.ImageIdx)
			} else if sh.IsText {
				desc = "TEXT"
			} else {
				desc = fmt.Sprintf("SHAPE(type=%d)", sh.ShapeType)
			}

			fillDesc := ""
			if sh.NoFill {
				fillDesc = "noFill"
			} else if sh.FillColor != "" {
				fillDesc = fmt.Sprintf("fill=%s", sh.FillColor)
				if sh.FillOpacity >= 0 && sh.FillOpacity < 65536 {
					fillDesc += fmt.Sprintf("@%d%%", sh.FillOpacity*100/65536)
				}
			}

			lineDesc := ""
			if sh.NoLine {
				lineDesc = "noLine"
			} else if sh.LineColor != "" {
				lineDesc = fmt.Sprintf("line=%s", sh.LineColor)
				if sh.LineWidth > 0 {
					lineDesc += fmt.Sprintf(" w=%d", sh.LineWidth)
				}
			}

			textDesc := ""
			totalChars := 0
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					totalChars += len(r.Text)
				}
			}
			if totalChars > 0 {
				firstText := ""
				for _, p := range sh.Paragraphs {
					for _, r := range p.Runs {
						firstText += r.Text
						if len(firstText) > 40 {
							break
						}
					}
					if len(firstText) > 40 {
						break
					}
				}
				if len(firstText) > 40 {
					firstText = firstText[:40] + "..."
				}
				textDesc = fmt.Sprintf(" text=%q", firstText)
			}

			// Only print first 10 shapes per slide for brevity, unless it's a problem slide
			if j < 10 || i == 42 { // slide 43 is the known problem
				fmt.Printf("  [%d] %s pos=(%d,%d) size=(%d,%d) %s %s%s\n",
					j, desc, sh.Left, sh.Top, sh.Width, sh.Height, fillDesc, lineDesc, textDesc)
			}
		}
		if len(shapes) > 10 && i != 42 {
			fmt.Printf("  ... (%d more shapes)\n", len(shapes)-10)
		}

		// Check for text content issues
		for _, sh := range shapes {
			for pi, p := range sh.Paragraphs {
				for ri, r := range p.Runs {
					if r.FontSize == 0 {
						fmt.Printf("  WARNING: slide %d shape has run with fontSize=0 (para=%d run=%d text=%q)\n", i+1, pi, ri, r.Text[:min(20, len(r.Text))])
					}
				}
			}
		}
	}

	// Now check the PPTX output
	fmt.Printf("\n\n=== PPTX Analysis ===\n")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Println("Error opening PPTX:", err)
		return
	}
	defer zr.Close()

	slideCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideCount++
		}
	}
	fmt.Printf("PPTX slides: %d\n", slideCount)

	// Check for common issues in PPTX
	issues := 0
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") || strings.Contains(f.Name, "_rels") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			continue
		}
		buf := make([]byte, 1024*1024)
		n, _ := rc.Read(buf)
		content := string(buf[:n])
		rc.Close()

		// Check for sz="0"
		if strings.Contains(content, `sz="0"`) {
			fmt.Printf("  ISSUE: %s contains sz=\"0\"\n", f.Name)
			issues++
		}

		// Check for missing showMasterSp
		if !strings.Contains(content, `showMasterSp=`) {
			fmt.Printf("  ISSUE: %s missing showMasterSp\n", f.Name)
			issues++
		}

		// Check for clrMapOvr
		if !strings.Contains(content, `clrMapOvr`) {
			fmt.Printf("  ISSUE: %s missing clrMapOvr\n", f.Name)
			issues++
		}
	}

	if issues == 0 {
		fmt.Println("No structural issues found in PPTX")
	}

	// Check specific slides for content
	fmt.Printf("\n=== Slide-by-slide text comparison (first 5 slides) ===\n")
	for i := 0; i < min(5, len(slides)); i++ {
		shapes := slides[i].GetShapes()
		fmt.Printf("\nSlide %d PPT text:\n", i+1)
		for _, sh := range shapes {
			for _, p := range sh.Paragraphs {
				line := ""
				for _, r := range p.Runs {
					line += r.Text
				}
				if strings.TrimSpace(line) != "" {
					fmt.Printf("  %q\n", line)
				}
			}
		}
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
