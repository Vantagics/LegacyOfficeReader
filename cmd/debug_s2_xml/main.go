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
		if zf.Name != "ppt/slides/slide2.xml" {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Find shapes with "背景与挑战" or "产品定位"
		targets := []string{"背景与挑战", "产品定位及价值", "典型场景", "最佳客户"}
		for _, target := range targets {
			idx := strings.Index(content, target)
			if idx < 0 {
				continue
			}
			// Find the enclosing <p:sp> 
			spStart := strings.LastIndex(content[:idx], "<p:sp>")
			spEnd := strings.Index(content[idx:], "</p:sp>")
			if spStart >= 0 && spEnd >= 0 {
				xml := content[spStart : idx+spEnd+7]
				// Just show the spPr part
				spPrStart := strings.Index(xml, "<p:spPr>")
				spPrEnd := strings.Index(xml, "</p:spPr>")
				if spPrStart >= 0 && spPrEnd >= 0 {
					fmt.Printf("=== %s ===\n%s\n\n", target, xml[spPrStart:spPrEnd+9])
				}
			}
		}

		// Also show the "1" shape
		idx := strings.Index(content, ">1<")
		if idx >= 0 {
			spStart := strings.LastIndex(content[:idx], "<p:sp>")
			spEnd := strings.Index(content[idx:], "</p:sp>")
			if spStart >= 0 && spEnd >= 0 {
				xml := content[spStart : idx+spEnd+7]
				spPrStart := strings.Index(xml, "<p:spPr>")
				spPrEnd := strings.Index(xml, "</p:spPr>")
				if spPrStart >= 0 && spPrEnd >= 0 {
					fmt.Printf("=== Number 1 shape ===\n%s\n\n", xml[spPrStart:spPrEnd+9])
				}
			}
		}
	}
}
