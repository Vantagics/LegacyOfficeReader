package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	re := regexp.MustCompile(`slideLayout(\d+)\.xml`)

	layoutUsage := make(map[string]int)
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/_rels/slide") && strings.HasSuffix(f.Name, ".xml.rels") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			m := re.FindStringSubmatch(string(data))
			if len(m) > 1 {
				layoutUsage[m[1]]++
			}
		}
	}

	fmt.Println("Layout usage:")
	for k, v := range layoutUsage {
		fmt.Printf("  Layout %s: %d slides\n", k, v)
	}

	// Check a specific slide with complex content (slide 5 has lots of shapes)
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide5.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Count shapes
			spCount := strings.Count(content, "<p:sp>")
			picCount := strings.Count(content, "<p:pic>")
			cxnCount := strings.Count(content, "<p:cxnSp>")
			fmt.Printf("\nSlide 5: sp=%d pic=%d cxn=%d total=%d\n", spCount, picCount, cxnCount, spCount+picCount+cxnCount)
		}
	}

	// Check slide 71 (the cloud deployment diagram)
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide71.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			spCount := strings.Count(content, "<p:sp>")
			picCount := strings.Count(content, "<p:pic>")
			cxnCount := strings.Count(content, "<p:cxnSp>")
			fmt.Printf("\nSlide 71: sp=%d pic=%d cxn=%d total=%d\n", spCount, picCount, cxnCount, spCount+picCount+cxnCount)

			// Check for color issues - look for empty solidFill
			emptyFill := strings.Count(content, `val=""/>`)
			fmt.Printf("  Empty color vals: %d\n", emptyFill)

			// Check for FFFFFF text on white background
			whiteFill := strings.Count(content, `val="FFFFFF"`)
			fmt.Printf("  FFFFFF vals: %d\n", whiteFill)
		}
	}
}
