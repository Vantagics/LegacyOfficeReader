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
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			content = strings.ReplaceAll(content, "><", ">\n<")
			fmt.Printf("\n=== %s (%d bytes) ===\n", f.Name, len(data))
			if len(content) > 3000 {
				content = content[:3000] + "\n... (truncated)"
			}
			fmt.Println(content)
		}
	}
}
