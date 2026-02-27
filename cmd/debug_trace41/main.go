package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
	"github.com/shakinm/xlsReader/convert/pptconv"
	"archive/zip"
	"encoding/xml"
	"io"
	"strings"
)

func main() {
	// First check what the PPT parser gives us
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	_ = p

	// Now check the PPTX output for slide 41
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

		// Find shapes with fill=000000
		content := string(data)
		idx := 0
		for {
			pos := strings.Index(content[idx:], `val="000000"`)
			if pos < 0 {
				break
			}
			absPos := idx + pos
			start := absPos - 300
			if start < 0 {
				start = 0
			}
			end := absPos + 300
			if end > len(content) {
				end = len(content)
			}
			fmt.Printf("@%d: ...%s...\n\n", absPos, content[start:end])
			idx = absPos + 12
		}
	}
	_ = pptconv.ConvertFile
	_ = xml.Unmarshal
}
