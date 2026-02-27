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
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Print each layout XML
	for i := 1; i <= 7; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", i))
		fmt.Printf("\n=== Layout %d ===\n%s\n", i, content)
	}

	// Print layout rels
	for i := 1; i <= 7; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slideLayouts/_rels/slideLayout%d.xml.rels", i))
		fmt.Printf("\n=== Layout %d rels ===\n%s\n", i, content)
	}

	// Print slide master
	content := readZipFile(zr, "ppt/slideMasters/slideMaster1.xml")
	fmt.Printf("\n=== Slide Master ===\n%s\n", content)

	// Print slide master rels
	content = readZipFile(zr, "ppt/slideMasters/_rels/slideMaster1.xml.rels")
	fmt.Printf("\n=== Slide Master rels ===\n%s\n", content)
}

func readZipFile(zr *zip.ReadCloser, name string) string {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			return string(data)
		}
	}
	return ""
}
