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
	path := "testfie/test_new.docx"
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

	// Split into paragraphs
	paraRe := regexp.MustCompile(`<w:p[ >].*?</w:p>|<w:p/>`)
	paras := paraRe.FindAllString(docXML, -1)

	fmt.Printf("Total paragraphs: %d\n\n", len(paras))

	// Find empty or near-empty paragraphs (no text content)
	textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
	imgRe := regexp.MustCompile(`<wp:inline|<wp:anchor|<w:drawing`)
	brRe := regexp.MustCompile(`<w:br `)

	emptyCount := 0
	for i, p := range paras {
		texts := textRe.FindAllStringSubmatch(p, -1)
		hasImg := imgRe.MatchString(p)
		hasBr := brRe.MatchString(p)

		allText := ""
		for _, t := range texts {
			allText += t[1]
		}
		allText = strings.TrimSpace(allText)

		if allText == "" && !hasImg && !hasBr {
			emptyCount++
			// Show context: what's around this paragraph
			snippet := p
			if len(snippet) > 200 {
				snippet = snippet[:200] + "..."
			}
			fmt.Printf("[%d] EMPTY: %s\n", i+1, snippet)
		}
	}
	fmt.Printf("\nTotal empty paragraphs: %d\n", emptyCount)
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
