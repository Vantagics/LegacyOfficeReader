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

	for _, f := range r.File {
		if f.Name == "ppt/slideLayouts/slideLayout4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Println(string(data))
		}
		if f.Name == "ppt/slideLayouts/_rels/slideLayout4.xml.rels" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Println("\n--- RELS ---")
			fmt.Println(string(data))
		}
	}
	
	// Also check what images are in the ppt/media folder
	fmt.Println("\n--- Media files ---")
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
		}
	}
}
