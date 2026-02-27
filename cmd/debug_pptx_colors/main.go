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

	// Check specific slides for white text issues
	checkSlides := []int{5, 9, 11, 12, 41, 64, 67, 70, 71}
	for _, sn := range checkSlides {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", sn)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				whiteCount := strings.Count(content, `val="FFFFFF"`)
				blackCount := strings.Count(content, `val="000000"`)
				fmt.Printf("Slide %d: FFFFFF=%d, 000000=%d, size=%d\n", sn, whiteCount, blackCount, len(data))

				// Check for FFD966 fills with text colors
				if strings.Contains(content, `val="FFD966"`) {
					// Find text near FFD966 fills
					idx := 0
					shown := 0
					for shown < 2 {
						pos := strings.Index(content[idx:], `val="FFD966"`)
						if pos < 0 {
							break
						}
						absPos := idx + pos
						start := absPos - 100
						if start < 0 { start = 0 }
						end := absPos + 400
						if end > len(content) { end = len(content) }
						fmt.Printf("  FFD966 context: ...%s...\n", content[start:end])
						idx = absPos + 1
						shown++
					}
				}
			}
		}
	}

	// Also check layout XMLs
	fmt.Println("\n=== Layout XMLs ===")
	for li := 1; li <= 7; li++ {
		fname := fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", li)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				hasGrad := strings.Contains(content, "gradFill")
				hasBg := strings.Contains(content, "<p:bg>")
				imgCount := strings.Count(content, "r:embed")
				fmt.Printf("Layout %d: size=%d, hasBg=%v, hasGrad=%v, imgRefs=%d\n", li, len(data), hasBg, hasGrad, imgCount)

				// Show first 500 chars
				show := content
				if len(show) > 800 {
					show = show[:800]
				}
				fmt.Printf("  Preview: %s\n\n", show)
			}
		}
	}
}
