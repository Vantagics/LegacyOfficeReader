package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check slides that use layout 4 (most common) for watermark image
	// Layout 4 watermark is imgIdx=13 -> image14.png
	for si := 8; si <= 12; si++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		for _, f := range zr.File {
			if f.Name != name {
				continue
			}
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			hasImage14 := strings.Contains(content, "rImg14")
			picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")

			// Check rels for image14
			relsName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", si)
			hasImage14Rel := false
			for _, rf := range zr.File {
				if rf.Name == relsName {
					rrc, _ := rf.Open()
					rdata, _ := io.ReadAll(rrc)
					rrc.Close()
					hasImage14Rel = strings.Contains(string(rdata), "image14")
				}
			}

			fmt.Printf("Slide %d: hasImage14=%v hasImage14Rel=%v pics=%d\n",
				si, hasImage14, hasImage14Rel, picCount)

			// Find the watermark image position in the slide XML
			if hasImage14 {
				idx := strings.Index(content, "rImg14")
				if idx >= 0 {
					// Get surrounding context
					start := idx - 300
					if start < 0 {
						start = 0
					}
					end := idx + 200
					if end > len(content) {
						end = len(content)
					}
					ctx := content[start:end]
					fmt.Printf("  Watermark context: ...%s...\n", ctx)
				}
			}
		}
	}

	// Check layout 4 for watermark
	fmt.Println("\n=== Layout 4 content ===")
	for _, f := range zr.File {
		if f.Name == "ppt/slideLayouts/slideLayout4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("%s\n", string(data))
		}
	}

	// Check layout 4 rels
	fmt.Println("\n=== Layout 4 rels ===")
	for _, f := range zr.File {
		if f.Name == "ppt/slideLayouts/_rels/slideLayout4.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("%s\n", string(data))
		}
	}
}
