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

	// Check all images
	fmt.Println("=== Images in PPTX ===")
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("  %s: %d bytes\n", f.Name, len(data))
		}
	}

	// Check slide 2 rels to verify image references
	for _, f := range zr.File {
		if f.Name == "ppt/slides/_rels/slide2.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== Slide 2 rels ===\n%s\n", string(data))
		}
	}

	// Check layout 4 rels
	for _, f := range zr.File {
		if f.Name == "ppt/slideLayouts/_rels/slideLayout4.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== Layout 4 rels ===\n%s\n", string(data))
		}
	}

	// Check slide master rels
	for _, f := range zr.File {
		if f.Name == "ppt/slideMasters/_rels/slideMaster1.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== Slide Master rels ===\n%s\n", string(data))
		}
	}

	// Check for any broken image references
	fmt.Println("\n=== Checking for broken image references ===")
	imageFiles := make(map[string]bool)
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			name := strings.TrimPrefix(f.Name, "ppt/media/")
			imageFiles[name] = true
		}
	}

	for _, f := range zr.File {
		if !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Find all image references
		for {
			idx := strings.Index(content, "../media/")
			if idx < 0 {
				break
			}
			content = content[idx+9:]
			end := strings.Index(content, `"`)
			if end < 0 {
				break
			}
			imgName := content[:end]
			if !imageFiles[imgName] {
				fmt.Printf("  BROKEN: %s references missing image %s\n", f.Name, imgName)
			}
			content = content[end:]
		}
	}
	fmt.Println("Done.")
}
