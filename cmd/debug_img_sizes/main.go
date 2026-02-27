package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"
)

func main() {
	path := "testfie/test_new7.docx"
	if len(os.Args) > 1 {
		path = os.Args[1]
	}

	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "failed to open: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	docXML := readZip(r, "word/document.xml")

	// Find all paragraphs with images and show their sizes
	paraRe := regexp.MustCompile(`<w:p[ >].*?</w:p>|<w:p/>`)
	paras := paraRe.FindAllString(docXML, -1)

	// EMU to cm conversion: 1 cm = 360000 EMU
	for i, p := range paras {
		if !strings.Contains(p, "<wp:inline") && !strings.Contains(p, "<wp:anchor") {
			continue
		}

		textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
		texts := textRe.FindAllStringSubmatch(p, -1)
		allText := ""
		for _, t := range texts {
			allText += t[1]
		}

		// Extract extent (cx, cy)
		extRe := regexp.MustCompile(`<wp:extent cx="(\d+)" cy="(\d+)"`)
		extMatch := extRe.FindStringSubmatch(p)
		cx, cy := "?", "?"
		cxCm, cyCm := 0.0, 0.0
		if extMatch != nil {
			cx = extMatch[1]
			cy = extMatch[2]
			var cxVal, cyVal int64
			fmt.Sscanf(cx, "%d", &cxVal)
			fmt.Sscanf(cy, "%d", &cyVal)
			cxCm = float64(cxVal) / 360000.0
			cyCm = float64(cyVal) / 360000.0
		}

		// Get image relationship
		relRe := regexp.MustCompile(`r:embed="([^"]+)"`)
		relMatch := relRe.FindStringSubmatch(p)
		relId := "?"
		if relMatch != nil {
			relId = relMatch[1]
		}

		// Context: show surrounding paragraph text
		prevText := ""
		if i > 0 {
			prevTexts := textRe.FindAllStringSubmatch(paras[i-1], -1)
			for _, t := range prevTexts {
				prevText += t[1]
			}
		}
		nextText := ""
		if i+1 < len(paras) {
			nextTexts := textRe.FindAllStringSubmatch(paras[i+1], -1)
			for _, t := range nextTexts {
				nextText += t[1]
			}
		}
		if len(prevText) > 40 {
			prevText = prevText[:40] + "..."
		}
		if len(nextText) > 40 {
			nextText = nextText[:40] + "..."
		}

		fmt.Printf("[%3d] IMAGE  %s=%s  size=%sx%s EMU (%.1fx%.1f cm)\n", i+1, relId, allText, cx, cy, cxCm, cyCm)
		fmt.Printf("      prev: %s\n", prevText)
		fmt.Printf("      next: %s\n\n", nextText)
	}
}

func readZip(r *zip.ReadCloser, name string) string {
	for _, f := range r.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			return string(data)
		}
	}
	return ""
}
