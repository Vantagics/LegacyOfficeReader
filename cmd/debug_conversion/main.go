package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	// Print full layout 3 XML
	for _, f := range r.File {
		if f.Name == "ppt/slideLayouts/slideLayout3.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("Layout 3:\n%s\n", string(data))
			break
		}
	}

	// Print full layout 4 XML
	for _, f := range r.File {
		if f.Name == "ppt/slideLayouts/slideLayout4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\nLayout 4:\n%s\n", string(data))
			break
		}
	}
}
