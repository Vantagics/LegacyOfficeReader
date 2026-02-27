package main

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"
)

func main() {
	f, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	// Check which layout each slide references
	slideLayoutMap := make(map[string]string)
	for _, file := range f.File {
		if !strings.HasPrefix(file.Name, "ppt/slides/_rels/slide") || !strings.HasSuffix(file.Name, ".xml.rels") {
			continue
		}
		rc, _ := file.Open()
		buf := make([]byte, file.UncompressedSize64)
		n, _ := rc.Read(buf)
		rc.Close()
		content := string(buf[:n])

		slideNum := file.Name[len("ppt/slides/_rels/slide") : len(file.Name)-len(".xml.rels")]
		// Find slideLayout reference
		idx := strings.Index(content, "slideLayout")
		if idx >= 0 {
			end := strings.Index(content[idx:], `"`)
			if end > 0 {
				layoutRef := content[idx : idx+end]
				slideLayoutMap[slideNum] = layoutRef
			}
		}
	}

	// Count slides per layout
	layoutCounts := make(map[string]int)
	for _, layout := range slideLayoutMap {
		layoutCounts[layout]++
	}

	fmt.Printf("=== Slide-to-Layout Mapping ===\n")
	for layout, count := range layoutCounts {
		fmt.Printf("  %s: %d slides\n", layout, count)
	}

	// Check layout content details
	fmt.Printf("\n=== Layout Details ===\n")
	for _, file := range f.File {
		if !strings.HasPrefix(file.Name, "ppt/slideLayouts/slideLayout") || !strings.HasSuffix(file.Name, ".xml") {
			continue
		}
		rc, _ := file.Open()
		buf := make([]byte, file.UncompressedSize64)
		n, _ := rc.Read(buf)
		rc.Close()
		content := string(buf[:n])

		hasBg := strings.Contains(content, "<p:bg>")
		hasBgImg := strings.Contains(content, "r:embed=") && strings.Contains(content, "<p:bg>")
		hasSolidFill := strings.Contains(content, "<a:solidFill>")
		shapeCount := strings.Count(content, "<p:sp>")
		picCount := strings.Count(content, "<p:pic>")
		connCount := strings.Count(content, "<p:cxnSp>")
		lineCount := strings.Count(content, "<a:ln>")

		fmt.Printf("  %s:\n", file.Name)
		fmt.Printf("    Background: %v (image=%v, solidFill=%v)\n", hasBg, hasBgImg, hasSolidFill)
		fmt.Printf("    Shapes: %d sp, %d pic, %d cxnSp, %d lines\n", shapeCount, picCount, connCount, lineCount)
		fmt.Printf("    Size: %d bytes\n", len(content))
	}

	// Check theme
	for _, file := range f.File {
		if file.Name == "ppt/theme/theme1.xml" {
			rc, _ := file.Open()
			buf := make([]byte, file.UncompressedSize64)
			n, _ := rc.Read(buf)
			rc.Close()
			content := string(buf[:n])

			fmt.Printf("\n=== Theme ===\n")
			fmt.Printf("  Size: %d bytes\n", len(content))
			// Check for dk1 color
			if idx := strings.Index(content, "dk1"); idx >= 0 {
				end := idx + 200
				if end > len(content) {
					end = len(content)
				}
				snippet := content[idx:end]
				fmt.Printf("  dk1 section: %s...\n", snippet)
			}
		}
	}
}
