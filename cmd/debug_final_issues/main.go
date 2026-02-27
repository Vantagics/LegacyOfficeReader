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
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()
	images := p.GetImages()
	w, h := p.GetSlideSize()

	fmt.Printf("PPT: %d slides, %d images, size=%dx%d\n", len(slides), len(images), w, h)

	// Check watermark image details
	fmt.Println("\n=== Watermark Image Analysis ===")
	// Layout 4 watermark is imgIdx=13
	if len(images) > 13 {
		img := images[13]
		fmt.Printf("Image 13: format=%d size=%d bytes\n", img.Format, len(img.Data))
	}

	// Check slide 8 (uses layout 4) - compare PPT shapes with PPTX output
	fmt.Println("\n=== Slide 8 PPT shapes ===")
	if len(slides) >= 8 {
		s := slides[7]
		shapes := s.GetShapes()
		ref := s.GetMasterRef()
		fmt.Printf("MasterRef=%d, Shapes=%d\n", ref, len(shapes))
		for si, sh := range shapes {
			if sh.IsText && len(sh.Paragraphs) > 0 {
				var text string
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						text += run.Text
					}
				}
				if len(text) > 50 {
					text = text[:50] + "..."
				}
				fmt.Printf("  [%d] TEXT type=%d pos=(%d,%d) sz=(%d,%d) fill=%s color=%s: %q\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor,
					getFirstColor(sh), text)
			} else if sh.IsImage {
				fmt.Printf("  [%d] IMAGE type=%d pos=(%d,%d) sz=(%d,%d) imgIdx=%d\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.ImageIdx)
			} else {
				fmt.Printf("  [%d] type=%d pos=(%d,%d) sz=(%d,%d) fill=%s line=%s\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.LineColor)
			}
		}
	}

	// Check slide 9 (architecture diagram) - lots of small shapes
	fmt.Println("\n=== Slide 9 PPT shapes (first 30) ===")
	if len(slides) >= 9 {
		s := slides[8]
		shapes := s.GetShapes()
		ref := s.GetMasterRef()
		m := masters[ref]
		fmt.Printf("MasterRef=%d, Shapes=%d, Scheme=%v\n", ref, len(shapes), m.ColorScheme)
		for si, sh := range shapes {
			if si >= 30 {
				fmt.Printf("  ... and %d more shapes\n", len(shapes)-30)
				break
			}
			if sh.IsText && len(sh.Paragraphs) > 0 {
				var text string
				var color string
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						text += run.Text
						if color == "" {
							color = run.Color
						}
					}
				}
				if len(text) > 40 {
					text = text[:40] + "..."
				}
				fmt.Printf("  [%d] TEXT type=%d pos=(%d,%d) sz=(%d,%d) fill=%s noFill=%v color=%s: %q\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.NoFill, color, text)
			} else if sh.IsImage {
				fmt.Printf("  [%d] IMAGE pos=(%d,%d) sz=(%d,%d) imgIdx=%d\n",
					si, sh.Left, sh.Top, sh.Width, sh.Height, sh.ImageIdx)
			} else {
				fmt.Printf("  [%d] type=%d pos=(%d,%d) sz=(%d,%d) fill=%s line=%s\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.LineColor)
			}
		}
	}

	// Check PPTX slide 9 for comparison
	fmt.Println("\n=== PPTX Slide 9 shape count ===")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide9.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			spCount := strings.Count(content, "<p:sp>")
			picCount := strings.Count(content, "<p:pic>")
			cxnCount := strings.Count(content, "<p:cxnSp>")
			fmt.Printf("PPTX Slide 9: sp=%d pic=%d cxn=%d total=%d\n",
				spCount, picCount, cxnCount, spCount+picCount+cxnCount)
		}
	}

	// Check slide 5 - the timeline slide with parallelogram shapes
	fmt.Println("\n=== Slide 5 PPT shapes ===")
	if len(slides) >= 5 {
		s := slides[4]
		shapes := s.GetShapes()
		ref := s.GetMasterRef()
		fmt.Printf("MasterRef=%d, Shapes=%d\n", ref, len(shapes))
		for si, sh := range shapes {
			if sh.IsText && len(sh.Paragraphs) > 0 {
				var text string
				var color string
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						text += run.Text
						if color == "" {
							color = run.Color
						}
					}
				}
				if len(text) > 60 {
					text = text[:60] + "..."
				}
				fmt.Printf("  [%d] TEXT type=%d pos=(%d,%d) sz=(%d,%d) fill=%s noFill=%v color=%s: %q\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.NoFill, color, text)
			} else if sh.IsImage {
				fmt.Printf("  [%d] IMAGE pos=(%d,%d) sz=(%d,%d) imgIdx=%d\n",
					si, sh.Left, sh.Top, sh.Width, sh.Height, sh.ImageIdx)
			} else {
				fmt.Printf("  [%d] type=%d pos=(%d,%d) sz=(%d,%d) fill=%s line=%s lineW=%d\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.LineColor, sh.LineWidth)
			}
		}
	}

	_ = h
}

func getFirstColor(sh ppt.ShapeFormatting) string {
	for _, para := range sh.Paragraphs {
		for _, run := range para.Runs {
			if run.Color != "" {
				return run.Color
			}
		}
	}
	return ""
}
