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

	// Check layouts 3 and 4 for gradient fill
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			rc, _ := f.Open()
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			rc.Close()
			content := buf.String()
			hasGrad := strings.Contains(content, "gradFill")
			hasTitleBg := strings.Contains(content, "Title Background")
			fmt.Printf("%s: hasGradient=%v, hasTitleBg=%v\n", f.Name, hasGrad, hasTitleBg)
			if hasGrad {
				// Print the gradient section
				idx := strings.Index(content, "gradFill")
				if idx >= 0 {
					start := idx - 100
					if start < 0 {
						start = 0
					}
					end := idx + 300
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("  Gradient context: ...%s...\n", content[start:end])
				}
			}
		}
	}
}
