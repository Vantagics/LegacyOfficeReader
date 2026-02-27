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

	slideCount := 0
	layoutCount := 0
	imageCount := 0
	var issues []string

	for _, file := range f.File {
		if strings.HasPrefix(file.Name, "ppt/slides/slide") && strings.HasSuffix(file.Name, ".xml") {
			slideCount++
			rc, err := file.Open()
			if err != nil {
				continue
			}
			buf := make([]byte, file.UncompressedSize64)
			n, _ := rc.Read(buf)
			rc.Close()
			content := string(buf[:n])

			// Check for sz="0" (missing font sizes)
			if strings.Contains(content, `sz="0"`) {
				issues = append(issues, fmt.Sprintf("%s: has sz=\"0\" (missing font size)", file.Name))
			}
			// Check for empty color values
			if strings.Contains(content, `val=""`) {
				issues = append(issues, fmt.Sprintf("%s: has val=\"\" (empty attribute)", file.Name))
			}
			// Check for scheme color leaks (colors that look like scheme refs)
			// Scheme colors in PPT have 0x08 prefix, which would show as "00XXXX" with wrong bytes
		}
		if strings.HasPrefix(file.Name, "ppt/slideLayouts/") && strings.HasSuffix(file.Name, ".xml") {
			layoutCount++
		}
		if strings.HasPrefix(file.Name, "ppt/media/") {
			imageCount++
		}
	}

	fmt.Printf("=== PPTX Quality Report ===\n")
	fmt.Printf("Slides: %d\n", slideCount)
	fmt.Printf("Layouts: %d\n", layoutCount)
	fmt.Printf("Images: %d\n", imageCount)
	fmt.Printf("Issues: %d\n", len(issues))
	for _, issue := range issues {
		fmt.Printf("  - %s\n", issue)
	}

	// Check a sample slide for color distribution
	fmt.Printf("\n=== Color Analysis (first 5 slides) ===\n")
	for _, file := range f.File {
		if !strings.HasPrefix(file.Name, "ppt/slides/slide") || !strings.HasSuffix(file.Name, ".xml") {
			continue
		}
		num := file.Name[len("ppt/slides/slide") : len(file.Name)-4]
		if num > "5" && len(num) > 1 {
			continue
		}
		rc, _ := file.Open()
		buf := make([]byte, file.UncompressedSize64)
		n, _ := rc.Read(buf)
		rc.Close()
		content := string(buf[:n])

		// Count color references
		colorCount := strings.Count(content, "<a:srgbClr")
		noFillCount := strings.Count(content, "<a:noFill/>")
		solidFillCount := strings.Count(content, "<a:solidFill>")
		fmt.Printf("  %s: %d colors, %d solidFills, %d noFills\n", file.Name, colorCount, solidFillCount, noFillCount)
	}

	// Check layout backgrounds
	fmt.Printf("\n=== Layout Analysis ===\n")
	for _, file := range f.File {
		if !strings.HasPrefix(file.Name, "ppt/slideLayouts/") || !strings.HasSuffix(file.Name, ".xml") {
			continue
		}
		rc, _ := file.Open()
		buf := make([]byte, file.UncompressedSize64)
		n, _ := rc.Read(buf)
		rc.Close()
		content := string(buf[:n])

		hasBg := strings.Contains(content, "<p:bg>")
		shapeCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:pic>") + strings.Count(content, "<p:cxnSp>")
		imgCount := strings.Count(content, "<p:pic>")
		fmt.Printf("  %s: bg=%v, shapes=%d, images=%d\n", file.Name, hasBg, shapeCount, imgCount)
	}

	if len(issues) == 0 {
		fmt.Printf("\n✓ No quality issues found\n")
	}
}
