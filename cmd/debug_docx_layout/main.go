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

	textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
	imgRe := regexp.MustCompile(`<wp:inline|<wp:anchor|<w:drawing`)
	brRe := regexp.MustCompile(`<w:br `)
	pageBreakRe := regexp.MustCompile(`w:type="page"`)
	headingRe := regexp.MustCompile(`<w:pStyle w:val="Heading(\d)"`)
	tocRe := regexp.MustCompile(`<w:pStyle w:val="TOC`)

	fmt.Printf("Total paragraphs: %d\n\n", len(paras))

	for i, p := range paras {
		texts := textRe.FindAllStringSubmatch(p, -1)
		hasImg := imgRe.MatchString(p)
		hasBr := brRe.MatchString(p)
		hasPageBreak := pageBreakRe.MatchString(p)
		headingMatch := headingRe.FindStringSubmatch(p)
		isTOC := tocRe.MatchString(p)

		allText := ""
		for _, t := range texts {
			allText += t[1]
		}
		allText = strings.TrimSpace(allText)

		// Determine type
		pType := "TEXT"
		if allText == "" && !hasImg && !hasBr {
			pType = "EMPTY"
		}
		if hasImg {
			pType = "IMAGE"
		}
		if hasBr && hasPageBreak {
			pType = "PAGEBREAK"
		}
		if len(headingMatch) > 0 {
			pType = fmt.Sprintf("H%s", headingMatch[1])
		}
		if isTOC {
			pType = "TOC"
		}
		if strings.Contains(p, "fldChar") {
			pType = "FIELD"
		}

		// Truncate text for display
		displayText := allText
		if len(displayText) > 60 {
			displayText = displayText[:60] + "..."
		}

		fmt.Printf("[%3d] %-10s %s\n", i+1, pType, displayText)
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
