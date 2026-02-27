package main

import (
	"archive/zip"
	"fmt"
	"io"

	"strings"
)

func main() {
	// Check reference PPTX
	zr, err := zip.OpenReader("testfie/reference.pptx")
	if err != nil {
		fmt.Printf("No reference.pptx: %v\n", err)
		return
	}
	defer zr.Close()

	fmt.Println("=== Reference PPTX Structure ===")
	for _, f := range zr.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Read slide1 from reference
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\nReference slide1.xml:\n%s\n", string(data))
		}
	}

	// Now check our output
	fmt.Println("\n=== Our PPTX Structure ===")
	zr2, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Printf("ERROR: %v\n", err)
		return
	}
	defer zr2.Close()

	for _, f := range zr2.File {
		if !strings.HasPrefix(f.Name, "ppt/media/") {
			fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
		}
	}

	// Count media files
	mediaCount := 0
	for _, f := range zr2.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaCount++
		}
	}
	fmt.Printf("  ppt/media/ files: %d\n", mediaCount)

	// Check a few slide XMLs for common issues
	issues := 0
	for _, f := range zr2.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Check for sz="0" (missing font size)
		if strings.Contains(content, `sz="0"`) {
			fmt.Printf("ISSUE: %s contains sz=\"0\"\n", f.Name)
			issues++
		}
		// Check for empty color
		if strings.Contains(content, `val=""/>`) {
			fmt.Printf("ISSUE: %s contains empty color val\n", f.Name)
			issues++
		}
	}
	if issues == 0 {
		fmt.Println("\nNo sz=0 or empty color issues found in slides.")
	}

	// Check layout XMLs
	for _, f := range zr2.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			hasShapes := strings.Contains(content, "<p:sp>") || strings.Contains(content, "<p:pic>")
			hasBg := strings.Contains(content, "<p:bg>")
			fmt.Printf("Layout %s: hasBg=%v hasShapes=%v size=%d\n", f.Name, hasBg, hasShapes, len(data))
		}
	}
}
