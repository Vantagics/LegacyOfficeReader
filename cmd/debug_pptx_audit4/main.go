package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

// Quick audit: compare PPT source shapes with generated PPTX slide content
func main() {
	// Parse source PPT
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "PPT open error: %v\n", err)
		os.Exit(1)
	}
	slides := p.GetSlides()
	fmt.Printf("PPT: %d slides\n", len(slides))

	// For each slide, show shape count and text summary
	for i, s := range slides {
		shapes := s.GetShapes()
		textShapes := 0
		imageShapes := 0
		otherShapes := 0
		totalChars := 0
		for _, sh := range shapes {
			if sh.IsImage && sh.ImageIdx >= 0 {
				imageShapes++
			} else if sh.IsText && len(sh.Paragraphs) > 0 {
				textShapes++
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						totalChars += len([]rune(run.Text))
					}
				}
			} else {
				otherShapes++
			}
		}
		bg := s.GetBackground()
		bgStr := "none"
		if bg.HasBackground {
			if bg.ImageIdx >= 0 {
				bgStr = fmt.Sprintf("img%d", bg.ImageIdx)
			} else if bg.FillColor != "" {
				bgStr = bg.FillColor
			}
		}
		fmt.Printf("  Slide %2d: %d shapes (text=%d img=%d other=%d) chars=%d bg=%s layout=%d masterRef=%d\n",
			i+1, len(shapes), textShapes, imageShapes, otherShapes, totalChars, bgStr, s.GetLayoutType(), s.GetMasterRef())
	}

	// Now inspect the generated PPTX
	fmt.Println("\n--- PPTX Inspection ---")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "PPTX open error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Count files by type
	slideFiles := 0
	layoutFiles := 0
	mediaFiles := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			slideFiles++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			layoutFiles++
		}
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaFiles++
		}
	}
	fmt.Printf("PPTX: %d slides, %d layouts, %d media files\n", slideFiles, layoutFiles, mediaFiles)

	// Check each slide for shape count and text content
	for i := 1; i <= slideFiles && i <= 10; i++ {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data := make([]byte, f.UncompressedSize64)
				rc.Read(data)
				rc.Close()
				content := string(data)

				// Count shapes
				spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
				picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
				cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
				hasBg := strings.Contains(content, "<p:bg>")
				showMaster := strings.Contains(content, `showMasterSp="1"`)

				fmt.Printf("  Slide %2d: sp=%d pic=%d cxn=%d bg=%v showMaster=%v\n",
					i, spCount, picCount, cxnCount, hasBg, showMaster)
				break
			}
		}
	}

	// Check first few layouts
	fmt.Println("\n--- Layout Details ---")
	for i := 1; i <= layoutFiles; i++ {
		fname := fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", i)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data := make([]byte, f.UncompressedSize64)
				rc.Read(data)
				rc.Close()
				content := string(data)

				spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
				picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
				cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
				hasBg := strings.Contains(content, "<p:bg>")
				showMaster := "n/a"
				if strings.Contains(content, `showMasterSp="0"`) {
					showMaster = "0"
				} else if strings.Contains(content, `showMasterSp="1"`) {
					showMaster = "1"
				}

				fmt.Printf("  Layout %d: sp=%d pic=%d cxn=%d bg=%v showMasterSp=%s\n",
					i, spCount, picCount, cxnCount, hasBg, showMaster)
				break
			}
		}
	}

	// Check a few specific slides for text content comparison
	fmt.Println("\n--- Text Content Spot Check (first 5 slides) ---")
	for i := 1; i <= 5 && i <= slideFiles; i++ {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data := make([]byte, f.UncompressedSize64)
				rc.Read(data)
				rc.Close()

				// Extract text content
				texts := extractTexts(data)
				preview := ""
				for _, t := range texts {
					t = strings.TrimSpace(t)
					if t != "" {
						if preview != "" {
							preview += " | "
						}
						if len([]rune(t)) > 30 {
							preview += string([]rune(t)[:30]) + "..."
						} else {
							preview += t
						}
					}
					if len(preview) > 120 {
						break
					}
				}
				fmt.Printf("  Slide %d: %s\n", i, preview)
				break
			}
		}
	}

	// Check slide master
	fmt.Println("\n--- Slide Master ---")
	for _, f := range zr.File {
		if f.Name == "ppt/slideMasters/slideMaster1.xml" {
			rc, _ := f.Open()
			data := make([]byte, f.UncompressedSize64)
			rc.Read(data)
			rc.Close()
			content := string(data)
			layoutRefs := strings.Count(content, "<p:sldLayoutId")
			fmt.Printf("  Layout refs: %d\n", layoutRefs)
			break
		}
	}
}

func extractTexts(data []byte) []string {
	var texts []string
	d := xml.NewDecoder(strings.NewReader(string(data)))
	inT := false
	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local == "t" {
				inT = true
			}
		case xml.EndElement:
			if t.Name.Local == "t" {
				inT = false
			}
		case xml.CharData:
			if inT {
				texts = append(texts, string(t))
			}
		}
	}
	return texts
}
