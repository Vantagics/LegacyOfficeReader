package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	// Check slide1 for color issues
	slides := []string{"ppt/slides/slide1.xml", "ppt/slides/slide2.xml", "ppt/slides/slide5.xml", "ppt/slides/slide6.xml"}
	for _, target := range slides {
		for _, f := range r.File {
			if f.Name == target {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Count color occurrences
				colors := map[string]int{}
				for _, c := range []string{"FFFFFF", "000000", "0C0D0E", "151A22", "656D78", "89919C", "FF0000"} {
					count := strings.Count(content, fmt.Sprintf(`val="%s"`, c))
					if count > 0 {
						colors[c] = count
					}
				}
				fmt.Printf("%s colors: %v\n", target, colors)

				// Check for sz="1800" (default) vs actual sizes
				sz1800 := strings.Count(content, `sz="1800"`)
				szTotal := strings.Count(content, `sz="`)
				fmt.Printf("  font sizes: %d total, %d are 1800 (default)\n", szTotal, sz1800)
			}
		}
	}
}
