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

	// Check slides with very large fonts
	for _, si := range []int{7, 9, 15, 20, 27} { // 0-indexed
		if si >= len(slides) {
			continue
		}
		s := slides[si]
		fmt.Printf("\n=== Slide %d ===\n", si+1)
		for j, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.FontSize >= 5000 && run.Text != "" {
						text := run.Text
						if len(text) > 40 {
							text = text[:40]
						}
						fmt.Printf("  Shape %d: sz=%d pos=(%d,%d) size=(%d,%d) %q\n",
							j, run.FontSize, sh.Left, sh.Top, sh.Width, sh.Height, text)
					}
				}
			}
		}
	}

	// Check PPTX for large fonts
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		panic(err)
	}
	defer zr.Close()

	szRe := regexp.MustCompile(`sz="(\d+)"`)
	for _, si := range []int{8, 10, 16, 21, 28} { // 1-indexed
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				matches := szRe.FindAllStringSubmatch(content, -1)
				for _, m := range matches {
					sz, _ := strconv.Atoi(m[1])
					if sz >= 5000 {
						// Find context around this match
						idx := strings.Index(content, m[0])
						start := idx - 100
						if start < 0 {
							start = 0
						}
						end := idx + 200
						if end > len(content) {
							end = len(content)
						}
						fmt.Printf("\nPPTX Slide %d large font sz=%d:\n  ...%s...\n", si, sz, content[start:end])
					}
				}
			}
		}
	}
}
