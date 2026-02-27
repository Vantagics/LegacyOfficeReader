package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	// Dump layout XML content
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			rc, _ := f.Open()
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			rc.Close()
			content := buf.String()
			fmt.Printf("\n=== %s ===\n", f.Name)
			// Print first 2000 chars
			if len(content) > 3000 {
				fmt.Println(content[:3000])
				fmt.Printf("... (%d more bytes)\n", len(content)-3000)
			} else {
				fmt.Println(content)
			}
		}
	}

	// Check a few slides for their layout reference
	for _, sn := range []int{1, 2, 4, 8, 10, 21, 58} {
		relsName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", sn)
		for _, f := range r.File {
			if f.Name == relsName {
				rc, _ := f.Open()
				var buf bytes.Buffer
				buf.ReadFrom(rc)
				rc.Close()
				fmt.Printf("\nSlide %d rels: %s\n", sn, buf.String())
			}
		}
	}

	// Check slide 1 XML for background handling
	for _, f := range r.File {
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			rc.Close()
			content := buf.String()
			fmt.Printf("\n=== slide1.xml ===\n")
			if len(content) > 3000 {
				fmt.Println(content[:3000])
			} else {
				fmt.Println(content)
			}
		}
	}

	// Check slide master
	for _, f := range r.File {
		if f.Name == "ppt/slideMasters/slideMaster1.xml" {
			rc, _ := f.Open()
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			rc.Close()
			fmt.Printf("\n=== slideMaster1.xml ===\n%s\n", buf.String())
		}
	}
}
