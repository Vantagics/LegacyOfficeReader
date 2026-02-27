package main

import (
	"archive/zip"
	"fmt"
	"os"
	"regexp"
	"strings"
)

func main() {
	f, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	layoutRe := regexp.MustCompile(`slideLayout(\d+)\.xml`)
	layoutCounts := make(map[string]int)
	var samples []string

	for _, file := range f.File {
		if !strings.HasPrefix(file.Name, "ppt/slides/_rels/slide") || !strings.HasSuffix(file.Name, ".xml.rels") {
			continue
		}
		rc, _ := file.Open()
		buf := make([]byte, file.UncompressedSize64)
		n, _ := rc.Read(buf)
		rc.Close()
		content := string(buf[:n])

		matches := layoutRe.FindStringSubmatch(content)
		if len(matches) > 1 {
			layoutNum := matches[1]
			layoutCounts[layoutNum]++
			if len(samples) < 5 || layoutNum != "4" {
				slideNum := file.Name[len("ppt/slides/_rels/slide") : len(file.Name)-len(".xml.rels")]
				samples = append(samples, fmt.Sprintf("slide%s -> layout%s", slideNum, layoutNum))
			}
		}
	}

	fmt.Printf("=== Layout Distribution ===\n")
	for layout, count := range layoutCounts {
		fmt.Printf("  Layout %s: %d slides\n", layout, count)
	}

	fmt.Printf("\n=== Sample Mappings ===\n")
	for _, s := range samples {
		fmt.Printf("  %s\n", s)
	}
}
