package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	f, err := os.Open("testfie/test.pptx")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	fi, _ := f.Stat()
	zr, err := zip.NewReader(f, fi.Size())
	if err != nil {
		panic(err)
	}

	// Read slide65.xml
	for _, zf := range zr.File {
		if zf.Name == "ppt/slides/slide65.xml" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find all font size references
			fmt.Println("=== Slide 65 XML (font sizes and colors) ===")
			lines := strings.Split(content, "<")
			for _, line := range lines {
				if strings.Contains(line, "sz=") || strings.Contains(line, "solidFill") || strings.Contains(line, "srgbClr") {
					fmt.Printf("<%s\n", line)
				}
			}
			return
		}
	}
	fmt.Println("slide65.xml not found")
}
