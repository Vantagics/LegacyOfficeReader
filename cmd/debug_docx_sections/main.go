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
	path := "testfie/test_new2.docx"
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

	// Check for sectPr elements
	sectRe := regexp.MustCompile(`<w:sectPr[^/]*>.*?</w:sectPr>`)
	sects := sectRe.FindAllString(docXML, -1)
	fmt.Printf("Section properties found: %d\n\n", len(sects))
	for i, s := range sects {
		snippet := s
		if len(snippet) > 500 {
			snippet = snippet[:500] + "..."
		}
		fmt.Printf("[%d] %s\n\n", i+1, snippet)
	}

	// Check for header/footer references
	hdrRe := regexp.MustCompile(`<w:headerReference[^/]*/?>`)
	ftrRe := regexp.MustCompile(`<w:footerReference[^/]*/?>`)
	hdrs := hdrRe.FindAllString(docXML, -1)
	ftrs := ftrRe.FindAllString(docXML, -1)
	fmt.Printf("\nHeader references: %d\n", len(hdrs))
	for _, h := range hdrs {
		fmt.Printf("  %s\n", h)
	}
	fmt.Printf("\nFooter references: %d\n", len(ftrs))
	for _, f := range ftrs {
		fmt.Printf("  %s\n", f)
	}

	// Check rels
	relsXML := readZip(r, "word/_rels/document.xml.rels")
	relRe := regexp.MustCompile(`<Relationship[^>]*>`)
	rels := relRe.FindAllString(relsXML, -1)
	fmt.Printf("\nDocument rels:\n")
	for _, rel := range rels {
		if strings.Contains(rel, "header") || strings.Contains(rel, "footer") ||
			strings.Contains(rel, "Header") || strings.Contains(rel, "Footer") {
			fmt.Printf("  %s\n", rel)
		}
	}

	// Check settings.xml for evenAndOddHeaders
	settingsXML := readZip(r, "word/settings.xml")
	if strings.Contains(settingsXML, "evenAndOddHeaders") {
		fmt.Printf("\nsettings.xml: evenAndOddHeaders FOUND\n")
	} else {
		fmt.Printf("\nsettings.xml: evenAndOddHeaders NOT FOUND\n")
	}

	// Check how many header/footer XML files exist
	fmt.Printf("\nHeader/Footer files in ZIP:\n")
	for _, f := range r.File {
		if strings.Contains(f.Name, "header") || strings.Contains(f.Name, "footer") {
			fmt.Printf("  %s\n", f.Name)
		}
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
