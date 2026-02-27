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

	// Check slide 4 and 5 for font sizes on FFD966 shapes
	for _, sn := range []int{4, 5, 6} {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", sn)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Find FFD966 shapes and check their text sizes
				idx := 0
				found := 0
				for found < 3 {
					pos := strings.Index(content[idx:], `val="FFD966"`)
					if pos < 0 {
						break
					}
					absPos := idx + pos
					// Find the enclosing <p:sp> ... </p:sp>
					spStart := strings.LastIndex(content[:absPos], "<p:sp>")
					if spStart < 0 {
						idx = absPos + 1
						continue
					}
					spEnd := strings.Index(content[absPos:], "</p:sp>")
					if spEnd < 0 {
						idx = absPos + 1
						continue
					}
					spEnd += absPos + len("</p:sp>")
					shape := content[spStart:spEnd]

					// Extract font sizes
					szIdx := 0
					var sizes []string
					for {
						szPos := strings.Index(shape[szIdx:], `sz="`)
						if szPos < 0 {
							break
						}
						szStart := szIdx + szPos + 4
						szEndPos := strings.Index(shape[szStart:], `"`)
						if szEndPos < 0 {
							break
						}
						sizes = append(sizes, shape[szStart:szStart+szEndPos])
						szIdx = szStart + szEndPos + 1
					}

					// Extract text
					var texts []string
					tIdx := 0
					for {
						tPos := strings.Index(shape[tIdx:], "<a:t>")
						if tPos < 0 {
							break
						}
						tStart := tIdx + tPos + 5
						tEnd := strings.Index(shape[tStart:], "</a:t>")
						if tEnd < 0 {
							break
						}
						texts = append(texts, shape[tStart:tStart+tEnd])
						tIdx = tStart + tEnd + 1
					}

					fmt.Printf("Slide %d FFD966 shape: sizes=%v texts=%v\n", sn, sizes, texts)
					idx = absPos + 1
					found++
				}
			}
		}
	}
}
