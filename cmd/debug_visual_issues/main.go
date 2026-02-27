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
	// Parse original PPT
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()
	w, h := p.GetSlideSize()

	// Check which masters have which images and how they map to layouts
	fmt.Println("=== Layout mapping ===")
	masterRefToLayoutIdx := make(map[uint32]int)
	var layoutRefs []uint32
	for _, s := range slides {
		ref := s.GetMasterRef()
		if _, ok := masterRefToLayoutIdx[ref]; !ok {
			idx := len(layoutRefs)
			masterRefToLayoutIdx[ref] = idx
			layoutRefs = append(layoutRefs, ref)
		}
	}

	for i, ref := range layoutRefs {
		m, ok := masters[ref]
		if !ok {
			fmt.Printf("Layout %d: masterRef=%d NOT FOUND\n", i+1, ref)
			continue
		}
		fmt.Printf("Layout %d: masterRef=%d bg=%v/%s/%d scheme=%v\n",
			i+1, ref, m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx, m.ColorScheme)

		// Count slides using this layout
		count := 0
		for _, s := range slides {
			if s.GetMasterRef() == ref {
				count++
			}
		}
		fmt.Printf("  Used by %d slides\n", count)

		// Show shapes
		for si, sh := range m.Shapes {
			isFullPage := sh.Width > int32(float64(w)*0.7) && sh.Height > int32(float64(h)*0.7)
			bottom := int64(sh.Top) + int64(sh.Height)
			isWatermark := sh.IsImage && sh.ImageIdx >= 0 && int64(sh.Top) > int64(h)/2 && bottom > int64(h)*3/4
			if sh.IsImage {
				fmt.Printf("  Shape[%d] IMAGE imgIdx=%d pos=(%d,%d) sz=(%d,%d) fullPage=%v watermark=%v\n",
					si, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height, isFullPage, isWatermark)
			} else if sh.IsText && len(sh.Paragraphs) > 0 {
				var text string
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						text += run.Text
					}
				}
				if len(text) > 60 {
					text = text[:60] + "..."
				}
				isPlaceholder := strings.Contains(text, "编辑母版") || strings.Contains(text, "www.") || text == "*"
				fmt.Printf("  Shape[%d] TEXT fill=%s placeholder=%v: %q\n",
					si, sh.FillColor, isPlaceholder, text)
			} else {
				fmt.Printf("  Shape[%d] type=%d pos=(%d,%d) sz=(%d,%d) line=%s\n",
					si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.LineColor)
			}
		}
	}

	// Check PPTX output for layout rendering
	fmt.Println("\n=== PPTX Layout Analysis ===")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
			picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
			cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
			hasBg := strings.Contains(content, "<p:bg>")
			hasGrad := strings.Contains(content, "gradFill")
			hasBlip := strings.Contains(content, "blipFill")
			showMaster := "?"
			if strings.Contains(content, `showMasterSp="1"`) {
				showMaster = "1"
			} else if strings.Contains(content, `showMasterSp="0"`) {
				showMaster = "0"
			}

			fmt.Printf("%s: sp=%d pic=%d cxn=%d bg=%v grad=%v blip=%v showMaster=%s\n",
				f.Name, spCount, picCount, cxnCount, hasBg, hasGrad, hasBlip, showMaster)
		}
	}

	// Check slide-layout relationships
	fmt.Println("\n=== Slide-Layout Relationships ===")
	for si := 1; si <= 10; si++ {
		name := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", si)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				fmt.Printf("Slide %d rels: %s\n", si, string(data))
			}
		}
	}

	// Check which slides have showMasterSp
	fmt.Println("\n=== Slide showMasterSp ===")
	for si := 1; si <= 10; si++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				showMaster := "?"
				if strings.Contains(content, `showMasterSp="1"`) {
					showMaster = "1"
				} else if strings.Contains(content, `showMasterSp="0"`) {
					showMaster = "0"
				}
				hasBg := strings.Contains(content, "<p:bg>")
				fmt.Printf("Slide %d: showMaster=%s hasBg=%v\n", si, showMaster, hasBg)
			}
		}
	}

	_ = h
}
