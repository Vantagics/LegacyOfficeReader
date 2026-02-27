package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	// Check the reference PPTX
	zr, err := zip.OpenReader("testfie/reference.pptx")
	if err != nil {
		fmt.Println("No reference.pptx found, checking test.pptx instead")
		return
	}
	defer zr.Close()

	fmt.Println("=== Reference PPTX structure ===")
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if len(content) > 500 {
				content = content[:500] + "..."
			}
			fmt.Printf("\n=== %s ===\n%s\n", f.Name, content)
		}
	}

	// Check slide 4 if it exists
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if len(content) > 2000 {
				content = content[:2000] + "..."
			}
			fmt.Printf("\n=== slide4.xml ===\n%s\n", content)
		}
	}

	// List all files
	fmt.Println("\n=== All files ===")
	for _, f := range zr.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}
}
