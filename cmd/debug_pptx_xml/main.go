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

	// Print slide1.xml content
	targets := []string{
		"ppt/slides/slide1.xml",
		"ppt/slides/slide2.xml",
		"ppt/slides/slide3.xml",
		"ppt/slideLayouts/slideLayout1.xml",
		"ppt/theme/theme1.xml",
		"ppt/slideMasters/slideMaster1.xml",
	}

	for _, target := range targets {
		for _, f := range r.File {
			if f.Name == target {
				fmt.Printf("\n=== %s ===\n", target)
				rc, err := f.Open()
				if err != nil {
					fmt.Fprintf(os.Stderr, "Error opening %s: %v\n", target, err)
					continue
				}
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				// Pretty print by adding newlines after >
				content = strings.ReplaceAll(content, "><", ">\n<")
				if len(content) > 5000 {
					content = content[:5000] + "\n... (truncated)"
				}
				fmt.Println(content)
			}
		}
	}

	// List all files
	fmt.Println("\n=== All files ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}
}
