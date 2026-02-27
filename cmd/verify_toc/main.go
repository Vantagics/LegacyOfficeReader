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

	// Count hyperlinks
	hlRe := regexp.MustCompile(`<w:hyperlink w:anchor="([^"]+)"`)
	hlMatches := hlRe.FindAllStringSubmatch(docXML, -1)
	fmt.Printf("=== Hyperlinks in document.xml (%d total) ===\n", len(hlMatches))
	for i, m := range hlMatches {
		fmt.Printf("  [%d] anchor=%s\n", i+1, m[1])
	}

	// Count bookmarks
	bmStartRe := regexp.MustCompile(`<w:bookmarkStart w:id="(\d+)" w:name="([^"]+)"`)
	bmMatches := bmStartRe.FindAllStringSubmatch(docXML, -1)
	fmt.Printf("\n=== Bookmarks in document.xml (%d total) ===\n", len(bmMatches))
	for i, m := range bmMatches {
		fmt.Printf("  [%d] id=%s name=%s\n", i+1, m[1], m[2])
	}

	// Check TOC field
	hasTOCBegin := strings.Contains(docXML, `TOC \o`)
	fmt.Printf("\n=== TOC field present: %v ===\n", hasTOCBegin)
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
