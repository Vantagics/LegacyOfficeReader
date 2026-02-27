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

	for _, zf := range r.File {
		if zf.Name != "ppt/slides/slide8.xml" {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Find the first <p:pic> element (watermark)
		picIdx := strings.Index(content, "<p:pic>")
		if picIdx < 0 {
			fmt.Println("No pic found")
			return
		}
		picEnd := strings.Index(content[picIdx:], "</p:pic>")
		if picEnd < 0 {
			fmt.Println("No pic end found")
			return
		}
		fmt.Println("=== Watermark pic XML ===")
		fmt.Println(content[picIdx : picIdx+picEnd+8])
	}
}
