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
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	masters := p.GetMasters()
	slides := p.GetSlides()

	// Check what master shapes look like (these become layout shapes)
	fmt.Printf("=== Master Slide Analysis ===\n")
	for ref, m := range masters {
		// Count how many slides use this master
		count := 0
		for _, s := range slides {
			if s.GetMasterRef() == ref {
				count++
			}
		}
		if count == 0 {
			continue
		}

		fmt.Printf("\nMaster %d (used by %d slides):\n", ref, count)
		fmt.Printf("  Background: has=%v fill=%s imgIdx=%d\n",
			m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
		fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
		fmt.Printf("  Shapes: %d\n", len(m.Shapes))

		for i, sh := range m.Shapes {
			textSnippet := ""
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t != "" {
						if len(t) > 40 {
							textSnippet = t[:40]
						} else {
							textSnippet = t
						}
						break
					}
				}
				if textSnippet != "" {
					break
				}
			}
			fmt.Printf("  [%d] type=%d img=%v imgIdx=%d fill=%s noFill=%v line=%s noLine=%v pos=(%d,%d) size=(%d,%d) %s\n",
				i, sh.ShapeType, sh.IsImage, sh.ImageIdx, sh.FillColor, sh.NoFill,
				sh.LineColor, sh.NoLine, sh.Left, sh.Top, sh.Width, sh.Height, textSnippet)
		}
	}

	// Check slide backgrounds
	fmt.Printf("\n=== Slide Background Analysis ===\n")
	bgCount := 0
	noBgCount := 0
	for i, s := range slides {
		bg := s.GetBackground()
		if bg.HasBackground {
			bgCount++
			fmt.Printf("Slide %d: bg fill=%s imgIdx=%d\n", i+1, bg.FillColor, bg.ImageIdx)
		} else {
			noBgCount++
		}
	}
	fmt.Printf("Slides with background: %d, without: %d\n", bgCount, noBgCount)

	// Check PPTX layout content
	fmt.Printf("\n=== PPTX Layout Content ===\n")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for i := 1; i <= 7; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", i))
		if content == "" {
			continue
		}
		// Check for background
		hasBg := strings.Contains(content, "<p:bg>")
		hasBlipBg := strings.Contains(content, "blipFill")
		hasSolidBg := strings.Contains(content, "solidFill") && strings.Contains(content, "<p:bg>")

		// Count shapes
		spCount := strings.Count(content, "<p:sp")
		picCount := strings.Count(content, "<p:pic")
		cxnCount := strings.Count(content, "<p:cxnSp")

		// Check for image references
		imgRefs := strings.Count(content, "r:embed=")

		fmt.Printf("Layout %d: bg=%v (blip=%v solid=%v) sp=%d pic=%d cxn=%d imgRefs=%d\n",
			i, hasBg, hasBlipBg, hasSolidBg, spCount, picCount, cxnCount, imgRefs)
	}

	// Check theme
	fmt.Printf("\n=== Theme Check ===\n")
	themeContent := readZipFile(zr, "ppt/theme/theme1.xml")
	if themeContent != "" {
		// Extract dk1 and lt1 colors
		dk1Idx := strings.Index(themeContent, "<a:dk1>")
		lt1Idx := strings.Index(themeContent, "<a:lt1>")
		if dk1Idx >= 0 {
			end := strings.Index(themeContent[dk1Idx:], "</a:dk1>")
			if end > 0 {
				fmt.Printf("dk1: %s\n", themeContent[dk1Idx:dk1Idx+end+8])
			}
		}
		if lt1Idx >= 0 {
			end := strings.Index(themeContent[lt1Idx:], "</a:lt1>")
			if end > 0 {
				fmt.Printf("lt1: %s\n", themeContent[lt1Idx:lt1Idx+end+8])
			}
		}
		// Check font
		if strings.Contains(themeContent, "微软雅黑") {
			fmt.Printf("Theme font: 微软雅黑 ✓\n")
		}
	}

	// Check slide 1 in detail (title slide)
	fmt.Printf("\n=== Slide 1 Detail ===\n")
	s1Content := readZipFile(zr, "ppt/slides/slide1.xml")
	fmt.Printf("Length: %d bytes\n", len(s1Content))
	// Print first 2000 chars
	if len(s1Content) > 2000 {
		fmt.Printf("%s...\n", s1Content[:2000])
	} else {
		fmt.Printf("%s\n", s1Content)
	}

	// Check slide 2 (typical content slide)
	fmt.Printf("\n=== Slide 2 Detail (first 3000 chars) ===\n")
	s2Content := readZipFile(zr, "ppt/slides/slide2.xml")
	if len(s2Content) > 3000 {
		fmt.Printf("%s...\n", s2Content[:3000])
	} else {
		fmt.Printf("%s\n", s2Content)
	}
}

func readZipFile(zr *zip.ReadCloser, name string) string {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			return string(data)
		}
	}
	return ""
}
