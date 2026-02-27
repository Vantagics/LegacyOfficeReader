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
		if !strings.HasPrefix(zf.Name, "ppt/slides/slide") && !strings.HasPrefix(zf.Name, "ppt/slideLayouts/slideLayout") {
			continue
		}
		if !strings.HasSuffix(zf.Name, ".xml") {
			continue
		}

		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		if strings.Contains(content, "srcRect") {
			// Count srcRect occurrences
			count := strings.Count(content, "srcRect")
			fmt.Printf("%s: %d srcRect elements\n", zf.Name, count)
		}
	}
}
