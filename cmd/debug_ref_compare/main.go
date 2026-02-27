package main

import (
	"archive/zip"
	"fmt"
	"io"

	"strings"
)

func main() {
	// Compare reference PPTX with generated PPTX
	fmt.Println("=== Reference PPTX (testfie/reference.pptx) ===")
	inspectPptx("testfie/reference.pptx")

	fmt.Println("\n=== Generated PPTX (testfie/test.pptx) ===")
	inspectPptx("testfie/test.pptx")
}

func inspectPptx(path string) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		fmt.Printf("  Cannot open: %v\n", err)
		return
	}
	defer zr.Close()

	slideCount := 0
	layoutCount := 0
	mediaCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideCount++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			layoutCount++
		}
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaCount++
		}
	}
	fmt.Printf("  Slides: %d, Layouts: %d, Media: %d\n", slideCount, layoutCount, mediaCount)

	// Check slide 1 structure
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			fmt.Printf("  Slide 1 size: %d bytes\n", len(data))
			fmt.Printf("  Slide 1 has bg: %v\n", strings.Contains(content, "<p:bg>"))
			fmt.Printf("  Slide 1 showMasterSp: %v\n", strings.Contains(content, `showMasterSp="1"`))
			break
		}
	}
}
