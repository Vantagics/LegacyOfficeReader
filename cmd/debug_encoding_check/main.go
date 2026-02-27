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
	// Check PPT text encoding
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	slides := p.GetSlides()

	fmt.Println("=== PPT Text (Go strings) ===")
	for i, s := range slides {
		if i >= 5 {
			break
		}
		shapes := s.GetShapes()
		for j, sh := range shapes {
			if !sh.IsText {
				continue
			}
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t != "" {
						fmt.Printf("  Slide %d Shape %d: font=%q text=%q\n", i+1, j, run.FontName, t)
						// Print hex of first few bytes
						if len(t) > 0 {
							bytes := []byte(t)
							if len(bytes) > 30 {
								bytes = bytes[:30]
							}
							fmt.Printf("    hex: %x\n", bytes)
						}
					}
				}
			}
		}
	}

	// Check PPTX output
	fmt.Println("\n=== PPTX Raw XML bytes ===")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "PPTX error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Find first <a:t> tag
			idx := strings.Index(content, "<a:t>")
			if idx >= 0 {
				end := strings.Index(content[idx:], "</a:t>")
				if end >= 0 {
					text := content[idx+5 : idx+end]
					fmt.Printf("  Slide 1 first text: %q\n", text)
					fmt.Printf("    hex: %x\n", []byte(text))
				}
			}
			break
		}
	}
}
