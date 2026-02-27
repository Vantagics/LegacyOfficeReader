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
	// Parse original PPT
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPT: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()
	fmt.Printf("PPT: %d slides\n", len(slides))

	// Show first few slides' key info
	for i := 0; i < len(slides) && i < 5; i++ {
		s := slides[i]
		shapes := s.GetShapes()
		bg := s.GetBackground()
		ref := s.GetMasterRef()
		fmt.Printf("\n--- PPT Slide %d (masterRef=%d, bg=%v/%s/%d) ---\n", i+1, ref, bg.HasBackground, bg.FillColor, bg.ImageIdx)
		fmt.Printf("  Shapes: %d\n", len(shapes))
		for si, sh := range shapes {
			if si > 15 {
				fmt.Printf("  ... and %d more shapes\n", len(shapes)-15)
				break
			}
			if sh.IsText && len(sh.Paragraphs) > 0 {
				var textParts []string
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						t := run.Text
						if len(t) > 60 {
							t = t[:60] + "..."
						}
						t = strings.ReplaceAll(t, "\n", "\\n")
						t = strings.ReplaceAll(t, "\r", "\\r")
						t = strings.ReplaceAll(t, "\x0b", "\\v")
						textParts = append(textParts, fmt.Sprintf("[%s sz=%d b=%v c=%s]%s", run.FontName, run.FontSize, run.Bold, run.Color, t))
					}
				}
				text := strings.Join(textParts, " | ")
				if len(text) > 200 {
					text = text[:200] + "..."
				}
				fmt.Printf("  Shape[%d] TEXT type=%d pos=(%d,%d) sz=(%d,%d) fill=%s noFill=%v: %s\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.NoFill, text)
			} else if sh.IsImage {
				fmt.Printf("  Shape[%d] IMAGE type=%d pos=(%d,%d) sz=(%d,%d) imgIdx=%d\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.ImageIdx)
			} else {
				fmt.Printf("  Shape[%d] OTHER type=%d pos=(%d,%d) sz=(%d,%d) fill=%s line=%s lineW=%d\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.LineColor, sh.LineWidth)
			}
		}
	}

	// Show master info
	for ref, m := range masters {
		fmt.Printf("\n--- Master ref=%d ---\n", ref)
		fmt.Printf("  BG: has=%v color=%s imgIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
		fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
		fmt.Printf("  Shapes: %d\n", len(m.Shapes))
		for si, sh := range m.Shapes {
			if sh.IsImage {
				fmt.Printf("  MasterShape[%d] IMAGE type=%d pos=(%d,%d) sz=(%d,%d) imgIdx=%d\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.ImageIdx)
			} else if sh.IsText && len(sh.Paragraphs) > 0 {
				var text string
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						text += run.Text
					}
				}
				if len(text) > 80 {
					text = text[:80] + "..."
				}
				fmt.Printf("  MasterShape[%d] TEXT type=%d pos=(%d,%d) sz=(%d,%d) fill=%s: %s\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, text)
			} else {
				fmt.Printf("  MasterShape[%d] type=%d pos=(%d,%d) sz=(%d,%d) fill=%s line=%s\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.LineColor)
			}
		}
	}

	// Now examine the generated PPTX
	fmt.Printf("\n\n=== PPTX Output Analysis ===\n")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Count slides and check first few
	slideFiles := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideFiles++
		}
	}
	fmt.Printf("PPTX slide files: %d\n", slideFiles)

	// Check first 3 slides XML
	for si := 1; si <= 3 && si <= slideFiles; si++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				if len(content) > 2000 {
					content = content[:2000] + "..."
				}
				fmt.Printf("\n--- PPTX slide%d.xml (first 2000 chars) ---\n%s\n", si, content)
			}
		}
	}

	// Check layout files
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if len(content) > 1500 {
				content = content[:1500] + "..."
			}
			fmt.Printf("\n--- %s (first 1500 chars) ---\n%s\n", f.Name, content)
		}
	}

	// Check theme
	for _, f := range zr.File {
		if strings.Contains(f.Name, "theme") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if len(content) > 1500 {
				content = content[:1500] + "..."
			}
			fmt.Printf("\n--- %s (first 1500 chars) ---\n%s\n", f.Name, content)
		}
	}

	_ = xml.Name{}
}
