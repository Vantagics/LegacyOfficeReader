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
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name != "ppt/slides/slide26.xml" {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Count total shapes
		spCount := strings.Count(content, "</p:sp>")
		picCount := strings.Count(content, "</p:pic>")
		fmt.Printf("Total sp: %d, pic: %d\n", spCount, picCount)

		// Count noFill vs solidFill
		noFillCount := strings.Count(content, "<a:noFill/>")
		solidFillCount := strings.Count(content, "<a:solidFill>")
		fmt.Printf("noFill: %d, solidFill: %d\n", noFillCount, solidFillCount)

		// Show first 3 shapes raw XML (truncated)
		parts := strings.Split(content, "<p:sp>")
		for i := 1; i < len(parts) && i <= 5; i++ {
			end := strings.Index(parts[i], "</p:sp>")
			if end < 0 {
				continue
			}
			xml := parts[i][:end]
			if len(xml) > 500 {
				xml = xml[:500] + "..."
			}
			fmt.Printf("\n--- Shape %d ---\n%s\n", i, xml)
		}
	}
}
