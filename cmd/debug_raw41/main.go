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
		if f.Name != "ppt/slides/slide41.xml" {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)
		// Find "数据库审计系统"
		pos := strings.Index(content, "数据库审计系统")
		if pos >= 0 {
			start := pos - 600
			if start < 0 {
				start = 0
			}
			end := pos + 200
			if end > len(content) {
				end = len(content)
			}
			fmt.Printf("Context around '数据库审计系统':\n%s\n", content[start:end])
		}
	}
}
