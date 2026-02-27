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
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Check a few slides for font and color issues
	checkSlides := []int{1, 2, 3, 5, 10, 20, 30, 50, 71}
	for _, num := range checkSlides {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", num)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Check for missing font
				noFont := strings.Count(content, `dirty="0"/>`) // self-closing rPr with no font
				hasFont := strings.Count(content, `typeface="微软雅黑"`)
				hasColor := strings.Count(content, `<a:solidFill>`)
				emptyRPr := strings.Count(content, `<a:rPr lang="zh-CN" altLang="en-US" dirty="0" sz="1800"/>`)

				fmt.Printf("Slide %d: fonts=%d noFont=%d colors=%d emptyRPr=%d len=%d\n",
					num, hasFont, noFont, hasColor, emptyRPr, len(content))

				// Show first 500 chars of text content
				if len(content) > 500 {
					// Find first <a:t> tag
					idx := strings.Index(content, "<a:t>")
					if idx > 0 {
						end := idx + 200
						if end > len(content) {
							end = len(content)
						}
						fmt.Printf("  First text: %s\n", content[idx:end])
					}
				}
				break
			}
		}
	}

	// Check layouts
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			fmt.Printf("Layout %s: len=%d hasBg=%v\n", f.Name, len(content),
				strings.Contains(content, "<p:bg>"))
		}
	}
}
