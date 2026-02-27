package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Parse PPT to check watermark/layout images
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	masters := p.GetMasters()
	images := p.GetImages()

	fmt.Printf("Total images: %d\n", len(images))

	// Check each master's shapes for image references
	for ref, m := range masters {
		fmt.Printf("\nMaster ref=%d: bg=%v (imgIdx=%d, fill=%s), shapes=%d, scheme=%v\n",
			ref, m.Background.HasBackground, m.Background.ImageIdx, m.Background.FillColor,
			len(m.Shapes), m.ColorScheme)
		for i, sh := range m.Shapes {
			if sh.IsImage && sh.ImageIdx >= 0 {
				ext := ""
				if sh.ImageIdx < len(images) {
					ext = fmt.Sprintf("format=%d, size=%d bytes", images[sh.ImageIdx].Format, len(images[sh.ImageIdx].Data))
				}
				fmt.Printf("  Shape %d: IMAGE idx=%d, pos=(%d,%d), size=(%d,%d), %s\n",
					i, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height, ext)
			}
		}
	}

	// Check the PPTX layout 4 (most used - 56 slides) for watermark
	fmt.Printf("\n\n=== Checking PPTX Layout 4 (watermark layout) ===\n")
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPTX: %v\n", err)
		return
	}
	defer r.Close()

	// Check layout 4 rels
	for _, f := range r.File {
		if f.Name == "ppt/slideLayouts/_rels/slideLayout4.xml.rels" {
			rc, _ := f.Open()
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			rc.Close()
			fmt.Printf("Layout 4 rels:\n%s\n", buf.String())
		}
	}

	// Check a slide that uses layout 4 (e.g., slide 8) for its content
	for _, sn := range []int{8, 11, 13} {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", sn)
		for _, f := range r.File {
			if f.Name == fname {
				rc, _ := f.Open()
				var buf bytes.Buffer
				buf.ReadFrom(rc)
				rc.Close()
				content := buf.String()
				// Check for title shape (first shape with white text on dark bg)
				titleIdx := strings.Index(content, "FFFFFF")
				if titleIdx >= 0 {
					start := titleIdx - 200
					if start < 0 {
						start = 0
					}
					end := titleIdx + 200
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("\nSlide %d - white text context:\n%s\n", sn, content[start:end])
				}
			}
		}
	}

	// Check if layout 4 has the watermark image properly positioned
	for _, f := range r.File {
		if f.Name == "ppt/slideLayouts/slideLayout4.xml" {
			rc, _ := f.Open()
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			rc.Close()
			content := buf.String()
			// Find all image references
			fmt.Printf("\nLayout 4 full XML:\n%s\n", content)
		}
	}
}
