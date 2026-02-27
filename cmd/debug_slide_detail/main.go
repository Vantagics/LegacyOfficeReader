package main

import (
	"archive/zip"
	"fmt"
	"io"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Check reference PPTX for slide 1 content
	zr, err := zip.OpenReader("testfie/reference.pptx")
	if err != nil {
		fmt.Println("No reference.pptx found, checking test.pptx")
	} else {
		defer zr.Close()
		for _, f := range zr.File {
			if f.Name == "ppt/slides/slide1.xml" {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				fmt.Printf("Reference slide1 size: %d\n", len(data))
				content := string(data)
				// Show first 2000 chars
				if len(content) > 2000 {
					content = content[:2000]
				}
				fmt.Println(content)
			}
		}
	}

	// Check our generated PPTX slide 1
	fmt.Println("\n=== Generated PPTX slide 1 ===")
	zr2, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer zr2.Close()
	for _, f := range zr2.File {
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("Generated slide1 size: %d\n", len(data))
			fmt.Println(string(data))
		}
	}

	// Check PPT slide 1 shapes in detail
	fmt.Println("\n=== PPT Slide 1 shapes ===")
	presentation, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	slides := presentation.GetSlides()
	if len(slides) > 0 {
		shapes := slides[0].GetShapes()
		bg := slides[0].GetBackground()
		fmt.Printf("Slide 1: %d shapes, bg=%v\n", len(shapes), bg)
		for i, sh := range shapes {
			fmt.Printf("  Shape %d: type=%d isText=%v isImage=%v pos=(%d,%d) size=(%d,%d)\n",
				i, sh.ShapeType, sh.IsText, sh.IsImage, sh.Left, sh.Top, sh.Width, sh.Height)
			fmt.Printf("    fill=%q noFill=%v line=%q noLine=%v lineW=%d\n",
				sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine, sh.LineWidth)
			fmt.Printf("    rot=%d flipH=%v flipV=%v\n", sh.Rotation, sh.FlipH, sh.FlipV)
			fmt.Printf("    margins: L=%d T=%d R=%d B=%d anchor=%d wrap=%d\n",
				sh.TextMarginLeft, sh.TextMarginTop, sh.TextMarginRight, sh.TextMarginBottom, sh.TextAnchor, sh.TextWordWrap)
			for pi, p := range sh.Paragraphs {
				fmt.Printf("    Para %d: align=%d indent=%d bullet=%v spacing=%d\n",
					pi, p.Alignment, p.IndentLevel, p.HasBullet, p.LineSpacing)
				for ri, r := range p.Runs {
					fmt.Printf("      Run %d: font=%q size=%d bold=%v italic=%v color=%q text=%q\n",
						ri, r.FontName, r.FontSize, r.Bold, r.Italic, r.Color, r.Text)
				}
			}
		}
	}

	// Also check slide 2 for the CONTENTS text with rotation
	fmt.Println("\n=== PPT Slide 2 shapes ===")
	if len(slides) > 1 {
		shapes := slides[1].GetShapes()
		for i, sh := range shapes {
			text := ""
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					text += r.Text
				}
			}
			if text != "" || sh.IsImage {
				fmt.Printf("  Shape %d: type=%d pos=(%d,%d) size=(%d,%d) rot=%d text=%q\n",
					i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.Rotation, text)
				if sh.IsImage {
					fmt.Printf("    IMAGE idx=%d\n", sh.ImageIdx)
				}
				for _, p := range sh.Paragraphs {
					for _, r := range p.Runs {
						fmt.Printf("    run: font=%q size=%d color=%q bold=%v\n", r.FontName, r.FontSize, r.Color, r.Bold)
					}
				}
			}
		}
	}

	// Check slide 11 (has dark fill shape)
	fmt.Println("\n=== PPT Slide 11 shapes with dark fill ===")
	if len(slides) > 10 {
		shapes := slides[10].GetShapes()
		for i, sh := range shapes {
			if sh.FillColor != "" && !sh.NoFill {
				text := ""
				for _, p := range sh.Paragraphs {
					for _, r := range p.Runs {
						text += r.Text
					}
				}
				fmt.Printf("  Shape %d: fill=%s opacity=%d text=%q\n", i, sh.FillColor, sh.FillOpacity, text)
				for _, p := range sh.Paragraphs {
					for _, r := range p.Runs {
						fmt.Printf("    run: size=%d color=%q\n", r.FontSize, r.Color)
					}
				}
			}
		}
	}

	// Check slide 41 (has many semi-transparent shapes)
	fmt.Println("\n=== PPT Slide 41 semi-transparent shapes ===")
	if len(slides) > 40 {
		shapes := slides[40].GetShapes()
		for i, sh := range shapes {
			if sh.FillOpacity > 0 && sh.FillOpacity < 65536 {
				text := ""
				for _, p := range sh.Paragraphs {
					for _, r := range p.Runs {
						text += r.Text
					}
				}
				fmt.Printf("  Shape %d: fill=%s opacity=%d (%d%%) text=%q\n",
					i, sh.FillColor, sh.FillOpacity, sh.FillOpacity*100/65536, text)
			}
		}
	}

	// Check for shapes with missing line info that should have lines
	fmt.Println("\n=== Shapes with lineWidth>0 but no lineColor ===")
	count := 0
	for i, s := range slides {
		for _, sh := range s.GetShapes() {
			if sh.LineWidth > 0 && sh.LineColor == "" && !sh.NoLine {
				count++
				if count <= 5 {
					text := ""
					for _, p := range sh.Paragraphs {
						for _, r := range p.Runs {
							text += r.Text
						}
					}
					fmt.Printf("  Slide %d: lineW=%d lineColor=%q text=%q\n", i+1, sh.LineWidth, sh.LineColor, text)
				}
			}
		}
	}
	fmt.Printf("Total: %d shapes with lineWidth>0 but no lineColor\n", count)

	// Check for shapes with explicit line color
	fmt.Println("\n=== Line color distribution ===")
	lineColors := make(map[string]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			if sh.LineColor != "" {
				lineColors[sh.LineColor]++
			}
		}
	}
	for c, n := range lineColors {
		fmt.Printf("  %s: %d\n", c, n)
	}
}
