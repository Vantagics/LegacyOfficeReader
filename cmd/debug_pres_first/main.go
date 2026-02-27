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
		if f.Name == "ppt/presentation.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			start := strings.Index(content, "<p:sldIdLst>")
			end := strings.Index(content, "</p:sldIdLst>")
			if start >= 0 && end >= 0 {
				section := content[start : end+len("</p:sldIdLst>")]
				entries := strings.Split(section, "<p:sldId")
				for i, e := range entries {
					if i == 0 {
						continue // skip the part before first sldId
					}
					if i <= 3 || i >= len(entries)-2 {
						fmt.Printf("  Entry %d: <p:sldId%s\n", i, strings.TrimRight(e, "/>"))
					}
				}
				fmt.Printf("  Total entries: %d\n", len(entries)-1)
			}
			break
		}
	}
}
