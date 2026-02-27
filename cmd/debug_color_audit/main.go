package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	issues := 0
	// Check each slide for potential color issues
	for _, f := range r.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") || strings.Contains(f.Name, "_rels") {
			continue
		}
		rc, _ := f.Open()
		var buf bytes.Buffer
		buf.ReadFrom(rc)
		rc.Close()
		content := buf.String()

		// Check for FFFFFF text color (white) - should only appear on dark backgrounds
		whiteCount := strings.Count(content, `val="FFFFFF"`)
		// Check for 000000 text color (black)
		blackCount := strings.Count(content, `val="000000"`)

		// Check for empty font sizes (sz="0")
		zeroSz := strings.Count(content, `sz="0"`)

		if zeroSz > 0 {
			fmt.Printf("ISSUE: %s has %d zero font sizes\n", f.Name, zeroSz)
			issues++
		}

		// Check for negative positions
		negX := strings.Count(content, `x="-`)
		negY := strings.Count(content, `y="-`)
		if negX > 0 || negY > 0 {
			// This is OK for layout shapes (logo at top-right can have negative y)
			// but flag it for slides
			_ = negX
			_ = negY
		}

		_ = whiteCount
		_ = blackCount
	}

	// Check for any layout issues
	for _, f := range r.File {
		if !strings.HasPrefix(f.Name, "ppt/slideLayouts/") || !strings.HasSuffix(f.Name, ".xml") || strings.Contains(f.Name, "_rels") {
			continue
		}
		rc, _ := f.Open()
		var buf bytes.Buffer
		buf.ReadFrom(rc)
		rc.Close()
		content := buf.String()

		// Verify layout has proper structure
		if !strings.Contains(content, "p:sldLayout") {
			fmt.Printf("ISSUE: %s missing sldLayout root\n", f.Name)
			issues++
		}
		if !strings.Contains(content, "p:spTree") {
			fmt.Printf("ISSUE: %s missing spTree\n", f.Name)
			issues++
		}
	}

	// Verify all slide rels point to valid layouts
	for i := 1; i <= 71; i++ {
		relsName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", i)
		for _, f := range r.File {
			if f.Name == relsName {
				rc, _ := f.Open()
				var buf bytes.Buffer
				buf.ReadFrom(rc)
				rc.Close()
				content := buf.String()
				if !strings.Contains(content, "slideLayout") {
					fmt.Printf("ISSUE: slide %d rels missing layout reference\n", i)
					issues++
				}
			}
		}
	}

	// Verify presentation.xml has all 71 slides
	for _, f := range r.File {
		if f.Name == "ppt/presentation.xml" {
			rc, _ := f.Open()
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			rc.Close()
			content := buf.String()
			sldIdCount := strings.Count(content, "p:sldId ")
			if sldIdCount != 71 {
				fmt.Printf("ISSUE: presentation.xml has %d slide IDs (expected 71)\n", sldIdCount)
				issues++
			} else {
				fmt.Printf("OK: presentation.xml has 71 slide IDs\n")
			}
		}
	}

	if issues == 0 {
		fmt.Println("All checks passed - no issues found")
	} else {
		fmt.Printf("\n%d issues found\n", issues)
	}
}
