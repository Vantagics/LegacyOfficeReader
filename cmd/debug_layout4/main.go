package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "ppt/slideLayouts/slideLayout4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("Layout 4:\n%s\n", string(data))
		}
		if f.Name == "ppt/slideLayouts/_rels/slideLayout4.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\nLayout 4 rels:\n%s\n", string(data))
		}
	}

	// Also check which master ref maps to layout 4
	fmt.Println("\n=== Checking master refs ===")
	// We need to check the PPT source
	// Layout 4 is the most used (56 slides)
}
