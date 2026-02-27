package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		panic(err)
	}
	defer zr.Close()

	// Print all layout XMLs
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== %s ===\n%s\n", f.Name, string(data))
		}
	}

	// Print slide master
	for _, f := range zr.File {
		if f.Name == "ppt/slideMasters/slideMaster1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== %s ===\n%s\n", f.Name, string(data))
		}
	}

	// Print theme
	for _, f := range zr.File {
		if f.Name == "ppt/theme/theme1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== %s ===\n%s\n", f.Name, string(data))
		}
	}

	// Print slide rels for slide 1 and 4
	for _, f := range zr.File {
		if f.Name == "ppt/slides/_rels/slide1.xml.rels" || f.Name == "ppt/slides/_rels/slide4.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== %s ===\n%s\n", f.Name, string(data))
		}
	}

	// Check which layout each slide uses
	fmt.Println("\n=== Slide-Layout Mapping ===")
	for i := 1; i <= 71; i++ {
		relName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", i)
		for _, f := range zr.File {
			if f.Name == relName {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				// Find slideLayout reference
				idx := strings.Index(content, "slideLayout")
				if idx >= 0 {
					end := strings.Index(content[idx:], `"`)
					if end > 0 {
						fmt.Printf("  Slide %d → %s\n", i, content[idx:idx+end])
					}
				}
			}
		}
	}
}
