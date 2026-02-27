package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name != "ppt/slides/slide61.xml" {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Find "跨境数据监测系统" and show its full shape
		idx := strings.Index(content, "跨境数据监测系统")
		if idx >= 0 {
			spStart := strings.LastIndex(content[:idx], "<p:sp>")
			spEnd := strings.Index(content[idx:], "</p:sp>")
			if spStart >= 0 && spEnd >= 0 {
				fmt.Printf("Shape: %s\n\n", content[spStart:idx+spEnd+len("</p:sp>")])
			}
		}

		// Find "数据库" at y=6067406
		idx = strings.Index(content, `y="6067406"`)
		if idx >= 0 {
			spStart := strings.LastIndex(content[:idx], "<p:sp>")
			spEnd := strings.Index(content[idx:], "</p:sp>")
			if spStart >= 0 && spEnd >= 0 {
				shape := content[spStart:idx+spEnd+len("</p:sp>")]
				if len(shape) < 1000 {
					fmt.Printf("Shape at y=6067406: %s\n\n", shape)
				}
			}
		}

		// Find "办公区"
		idx = strings.Index(content, "办公区")
		if idx >= 0 {
			spStart := strings.LastIndex(content[:idx], "<p:sp>")
			spEnd := strings.Index(content[idx:], "</p:sp>")
			if spStart >= 0 && spEnd >= 0 {
				fmt.Printf("办公区 shape: %s\n\n", content[spStart:idx+spEnd+len("</p:sp>")])
			}
		}
	}
}
