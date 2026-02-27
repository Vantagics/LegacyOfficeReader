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
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check all slides for remaining white text issues
	totalWhite := 0
	totalBlack := 0
	issueSlides := 0

	for sn := 1; sn <= 71; sn++ {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", sn)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				white := strings.Count(content, `val="FFFFFF"`)
				black := strings.Count(content, `val="000000"`)
				totalWhite += white
				totalBlack += black

				if white > 10 {
					issueSlides++
					fmt.Printf("Slide %d: FFFFFF=%d, 000000=%d\n", sn, white, black)
				}
			}
		}
	}
	fmt.Printf("\nTotal: FFFFFF=%d, 000000=%d, slides with >10 white=%d\n", totalWhite, totalBlack, issueSlides)

	// Specifically check slides that had issues before
	fmt.Println("\n=== Specific slide checks ===")
	for _, sn := range []int{5, 9, 11, 12, 27, 41, 61, 63, 67, 70} {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", sn)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				white := strings.Count(content, `val="FFFFFF"`)
				black := strings.Count(content, `val="000000"`)
				fmt.Printf("Slide %d: FFFFFF=%d, 000000=%d\n", sn, white, black)
			}
		}
	}

	// Check slide 9 for specific shapes that should now be dark
	fmt.Println("\n=== Slide 9 sample shapes ===")
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide9.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find "安全态势大屏" text
			idx := strings.Index(content, "安全态势大屏")
			if idx >= 0 {
				start := idx - 300
				if start < 0 { start = 0 }
				end := idx + 100
				if end > len(content) { end = len(content) }
				fmt.Printf("安全态势大屏 context: ...%s...\n", content[start:end])
			}

			// Find "客户价值" on slide 11
			break
		}
	}

	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide11.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			idx := strings.Index(content, "客户价值")
			if idx >= 0 {
				start := idx - 300
				if start < 0 { start = 0 }
				end := idx + 100
				if end > len(content) { end = len(content) }
				fmt.Printf("\n客户价值 context: ...%s...\n", content[start:end])
			}
			break
		}
	}

	// Check slide 63 for E9EBF5 shapes
	fmt.Println("\n=== Slide 63 E9EBF5 check ===")
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide63.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			white := strings.Count(content, `val="FFFFFF"`)
			black := strings.Count(content, `val="000000"`)
			fmt.Printf("FFFFFF=%d, 000000=%d\n", white, black)

			// Check first E9EBF5 shape
			idx := strings.Index(content, `val="E9EBF5"`)
			if idx >= 0 {
				start := idx - 100
				if start < 0 { start = 0 }
				end := idx + 400
				if end > len(content) { end = len(content) }
				fmt.Printf("Sample: ...%s...\n", content[start:end])
			}
			break
		}
	}
}
