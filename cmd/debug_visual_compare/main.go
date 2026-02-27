package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Parse the PPT source
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPT: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()
	images := p.GetImages()
	slideW, slideH := p.GetSlideSize()

	fmt.Printf("=== PPT Source Analysis ===\n")
	fmt.Printf("Slides: %d, Images: %d, Masters: %d\n", len(slides), len(images), len(masters))
	fmt.Printf("Slide size: %d x %d EMU (%.1f x %.1f inches)\n", slideW, slideH, float64(slideW)/914400, float64(slideH)/914400)

	// Analyze each slide
	for i, s := range slides {
		shapes := s.GetShapes()
		bg := s.GetBackground()
		ref := s.GetMasterRef()

		textShapes := 0
		imageShapes := 0
		connectorShapes := 0
		otherShapes := 0
		totalTextLen := 0

		for _, sh := range shapes {
			switch {
			case sh.IsImage && sh.ImageIdx >= 0:
				imageShapes++
			case isConnector(sh.ShapeType):
				connectorShapes++
			case sh.IsText || len(sh.Paragraphs) > 0:
				textShapes++
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						totalTextLen += len(run.Text)
					}
				}
			default:
				otherShapes++
			}
		}

		fmt.Printf("\nSlide %d: masterRef=%d, shapes=%d (text=%d, img=%d, conn=%d, other=%d), textLen=%d",
			i+1, ref, len(shapes), textShapes, imageShapes, connectorShapes, otherShapes, totalTextLen)
		if bg.HasBackground {
			if bg.ImageIdx >= 0 {
				fmt.Printf(", bg=image(%d)", bg.ImageIdx)
			} else if bg.FillColor != "" {
				fmt.Printf(", bg=#%s", bg.FillColor)
			}
		}
		fmt.Println()

		// Show first few shapes with details for problem slides
		if i < 5 || i == 3 || i == 8 || i == 12 || i == 40 {
			for j, sh := range shapes {
				if j >= 8 {
					fmt.Printf("    ... and %d more shapes\n", len(shapes)-8)
					break
				}
				fmt.Printf("    Shape %d: type=%d, pos=(%d,%d), size=(%d,%d)",
					j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height)
				if sh.IsImage {
					fmt.Printf(", IMAGE(idx=%d)", sh.ImageIdx)
				}
				if sh.FillColor != "" {
					fmt.Printf(", fill=#%s", sh.FillColor)
				}
				if sh.NoFill {
					fmt.Printf(", noFill")
				}
				if sh.LineColor != "" {
					fmt.Printf(", line=#%s(w=%d)", sh.LineColor, sh.LineWidth)
				}
				if len(sh.Paragraphs) > 0 {
					firstText := ""
					for _, para := range sh.Paragraphs {
						for _, run := range para.Runs {
							if run.Text != "" {
								firstText = run.Text
								break
							}
						}
						if firstText != "" {
							break
						}
					}
					if len(firstText) > 40 {
						firstText = firstText[:40] + "..."
					}
					fmt.Printf(", text=%q", firstText)
					// Show font info from first run
					for _, para := range sh.Paragraphs {
						for _, run := range para.Runs {
							if run.Text != "" {
								fmt.Printf(", font=%q sz=%d", run.FontName, run.FontSize)
								if run.Color != "" {
									fmt.Printf(" color=#%s", run.Color)
								}
								if run.Bold {
									fmt.Printf(" B")
								}
								break
							}
						}
						break
					}
				}
				fmt.Println()
			}
		}
	}

	// Now analyze the generated PPTX
	fmt.Printf("\n\n=== PPTX Output Analysis ===\n")
	analyzePPTX("testfie/test.pptx")
}

func isConnector(shapeType uint16) bool {
	switch shapeType {
	case 20, 32, 33, 34, 35, 36, 37, 38, 39, 40:
		return true
	}
	return false
}

type xmlNode struct {
	XMLName  xml.Name
	Attrs    []xml.Attr `xml:",any,attr"`
	Content  []byte     `xml:",chardata"`
	Children []xmlNode  `xml:",any"`
}

func analyzePPTX(path string) {
	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPTX: %v\n", err)
		return
	}
	defer r.Close()

	// Count files by type
	slideFiles := 0
	layoutFiles := 0
	mediaFiles := 0
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideFiles++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			layoutFiles++
		}
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaFiles++
		}
	}
	fmt.Printf("PPTX: %d slides, %d layouts, %d media files\n", slideFiles, layoutFiles, mediaFiles)

	// Analyze a few key slides
	checkSlides := []int{1, 2, 3, 4, 5, 9, 13, 41, 71}
	for _, sn := range checkSlides {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", sn)
		data := readZipFile(r, fname)
		if data == nil {
			continue
		}
		analyzeSlideXML(sn, data)
	}

	// Analyze layouts
	for li := 1; li <= layoutFiles; li++ {
		fname := fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", li)
		data := readZipFile(r, fname)
		if data == nil {
			continue
		}
		analyzeLayoutXML(li, data)
	}
}

func readZipFile(r *zip.ReadCloser, name string) []byte {
	for _, f := range r.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil
			}
			defer rc.Close()
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			return buf.Bytes()
		}
	}
	return nil
}

func analyzeSlideXML(slideNum int, data []byte) {
	content := string(data)

	// Count shapes
	spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
	picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
	cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
	hasBg := strings.Contains(content, "<p:bg>")
	hasShowMaster := strings.Contains(content, `showMasterSp="1"`)

	fmt.Printf("\nPPTX Slide %d: sp=%d, pic=%d, cxn=%d, bg=%v, showMaster=%v\n",
		slideNum, spCount, picCount, cxnCount, hasBg, hasShowMaster)
}

func analyzeLayoutXML(layoutNum int, data []byte) {
	content := string(data)

	spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
	picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
	cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
	hasBg := strings.Contains(content, "<p:bg>")
	hasBlipBg := strings.Contains(content, "blipFill") && hasBg

	fmt.Printf("Layout %d: sp=%d, pic=%d, cxn=%d, bg=%v (blip=%v)\n",
		layoutNum, spCount, picCount, cxnCount, hasBg, hasBlipBg)
}
