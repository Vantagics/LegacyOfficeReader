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
	// 1. Check source PPT master/layout watermark info
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()

	// Slide 8 master ref
	slide8 := slides[7]
	ref := slide8.GetMasterRef()
	fmt.Printf("Slide 8: masterRef=%d\n", ref)

	if m, ok := masters[ref]; ok {
		fmt.Printf("Master shapes: %d\n", len(m.Shapes))
		for i, sh := range m.Shapes {
			if sh.IsImage {
				fmt.Printf("  Shape %d: IMAGE idx=%d pos=(%d,%d) sz=(%d,%d)\n",
					i, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height)
			}
		}
	}

	// 2. Check output PPTX slide 8
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening pptx: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, zf := range r.File {
		if zf.Name != "ppt/slides/slide8.xml" {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		picCount := strings.Count(content, "<p:pic>")
		fmt.Printf("\nPPTX slide 8: %d <p:pic> elements\n", picCount)

		// Find all blip references
		idx := 0
		for {
			pos := strings.Index(content[idx:], `r:embed="`)
			if pos < 0 {
				break
			}
			start := idx + pos + len(`r:embed="`)
			end := strings.Index(content[start:], `"`)
			if end < 0 {
				break
			}
			relID := content[start : start+end]
			fmt.Printf("  blip embed: %s\n", relID)
			idx = start + end + 1
		}
	}

	// Check slide 8 rels
	for _, zf := range r.File {
		if zf.Name != "ppt/slides/_rels/slide8.xml.rels" {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		fmt.Printf("\nSlide 8 rels:\n%s\n", string(data))
	}
}
