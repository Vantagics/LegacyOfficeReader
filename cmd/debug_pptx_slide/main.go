package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	slideNums := []string{"13", "41", "71"}
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		for _, sn := range slideNums {
			if f.Name == "ppt/slides/slide"+sn+".xml" {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				// Find shapes with solidFill and text
				fmt.Printf("=== %s (len=%d) ===\n", f.Name, len(content))
				// Print sections around "FFFFFF" color references
				idx := 0
				for {
					pos := strings.Index(content[idx:], "FFFFFF")
					if pos < 0 {
						break
					}
					absPos := idx + pos
					start := absPos - 200
					if start < 0 {
						start = 0
					}
					end := absPos + 200
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("  @%d: ...%s...\n\n", absPos, content[start:end])
					idx = absPos + 6
				}
			}
		}
	}
}
