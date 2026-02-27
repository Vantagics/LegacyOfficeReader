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

	for _, f := range zr.File {
		if strings.Contains(f.Name, "slideLayout") || strings.Contains(f.Name, "slideMaster") || strings.Contains(f.Name, "theme") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if len(content) > 1500 {
				content = content[:1500] + "..."
			}
			fmt.Printf("=== %s ===\n%s\n\n", f.Name, content)
		}
	}
}
