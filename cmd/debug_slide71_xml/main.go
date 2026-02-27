package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test_compare.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide71.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Find shapes with "流量采集" or "VM" or "vSwitch" or "Hypervisor"
			keywords := []string{"Agent", "vSwitch", "Hypervisor", "Docker Engine", "VM"}
			for _, kw := range keywords {
				idx := strings.Index(content, kw)
				if idx >= 0 {
					start := idx - 800
					if start < 0 {
						start = 0
					}
					end := idx + 200
					if end > len(content) {
						end = len(content)
					}
					fmt.Printf("=== Context for '%s' ===\n%s\n\n", kw, content[start:end])
					break // just show first match
				}
			}
			// Also show the shape with "流量采集分发平台"
			idx := strings.Index(content, "AE5A21")
			if idx >= 0 {
				start := idx - 500
				if start < 0 { start = 0 }
				end := idx + 500
				if end > len(content) { end = len(content) }
				fmt.Printf("=== Context for AE5A21 ===\n%s\n\n", content[start:end])
			}
			// Show all srgbClr values
			for i := 0; i < len(content)-20; i++ {
				if content[i:i+8] == "srgbClr " {
					end := i + 40
					if end > len(content) { end = len(content) }
					fmt.Printf("srgbClr: %s\n", content[i:end])
				}
			}
		}
	}
}
