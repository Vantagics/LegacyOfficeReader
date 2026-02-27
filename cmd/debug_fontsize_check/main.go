package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
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

	// Check slides 4 and 5 for the yellow banner shapes with sz=0
	for _, si := range []int{3, 4} { // 0-indexed
		if si >= len(slides) {
			continue
		}
		s := slides[si]
		fmt.Printf("=== Slide %d ===\n", si+1)
		for j, sh := range s.GetShapes() {
			hasSz0 := false
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.FontSize == 0 && run.Text != "" {
						hasSz0 = true
						break
					}
				}
				if hasSz0 {
					break
				}
			}
			if hasSz0 {
				text := ""
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						text += run.Text
					}
					text += "|"
				}
				if len(text) > 80 {
					text = text[:80]
				}
				fmt.Printf("  Shape %d: type=%d pos=(%d,%d) size=(%d,%d) fill=%q paras=%d %q\n",
					j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, len(sh.Paragraphs), text)
			}
		}
	}

	// Now check what font sizes appear in the PPTX output for these slides
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		panic(err)
	}
	defer zr.Close()

	szRe := regexp.MustCompile(`sz="(\d+)"`)
	for _, si := range []int{4, 5} { // 1-indexed
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Find all font sizes
				matches := szRe.FindAllStringSubmatch(content, -1)
				szDist := make(map[int]int)
				for _, m := range matches {
					sz, _ := strconv.Atoi(m[1])
					szDist[sz]++
				}
				fmt.Printf("\nPPTX Slide %d font size distribution: %v\n", si, szDist)

				// Check for the yellow banner text
				if strings.Contains(content, "FFD966") {
					idx := strings.Index(content, "FFD966")
					start := idx - 200
					if start < 0 {
						start = 0
					}
					end := idx + 500
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("  Yellow banner context: ...%s...\n", content[start:end])
				}
			}
		}
	}
}
