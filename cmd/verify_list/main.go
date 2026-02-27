package main

import (
	"archive/zip"
	"fmt"
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

	// Read document.xml
	docXML := readZipFile(r, "word/document.xml")
	numXML := readZipFile(r, "word/numbering.xml")

	// Count numPr references
	numPrRe := regexp.MustCompile(`<w:numPr><w:ilvl w:val="(\d+)"/><w:numId w:val="(\d+)"/></w:numPr>`)
	matches := numPrRe.FindAllStringSubmatch(docXML, -1)
	fmt.Printf("=== List references in document.xml (%d total) ===\n", len(matches))
	for i, m := range matches {
		fmt.Printf("  [%d] ilvl=%s numId=%s\n", i+1, m[1], m[2])
	}

	// Count abstractNum and num definitions
	absCount := strings.Count(numXML, "<w:abstractNum ")
	numCount := strings.Count(numXML, "<w:num ")
	fmt.Printf("\n=== numbering.xml ===\n")
	fmt.Printf("  abstractNum definitions: %d\n", absCount)
	fmt.Printf("  num definitions: %d\n", numCount)

	// Show each abstractNum's format
	absRe := regexp.MustCompile(`<w:abstractNum w:abstractNumId="(\d+)">(.*?)</w:abstractNum>`)
	absMatches := absRe.FindAllStringSubmatch(numXML, -1)
	for _, m := range absMatches {
		isBullet := strings.Contains(m[2], `numFmt w:val="bullet"`)
		isDecimal := strings.Contains(m[2], `numFmt w:val="decimal"`)
		isRoman := strings.Contains(m[2], `Roman"`)
		isLetter := strings.Contains(m[2], `Letter"`)
		// Extract the actual numFmt value
		fmtRe := regexp.MustCompile(`numFmt w:val="([^"]+)"`)
		fmtMatch := fmtRe.FindStringSubmatch(m[2])
		fmtVal := ""
		if fmtMatch != nil {
			fmtVal = fmtMatch[1]
		}
		// Extract lvlText for level 0
		lvlTextRe := regexp.MustCompile(`<w:lvl w:ilvl="0">.*?<w:lvlText w:val="([^"]*)"`)
		lvlTextMatch := lvlTextRe.FindStringSubmatch(m[2])
		lvlText := ""
		if lvlTextMatch != nil {
			lvlText = lvlTextMatch[1]
		}
		fmt.Printf("  abstractNumId=%s: fmt=%s lvlText=%q bullet=%v decimal=%v roman=%v letter=%v\n", m[1], fmtVal, lvlText, isBullet, isDecimal, isRoman, isLetter)
	}

	// Show num -> abstractNum mapping
	numRe := regexp.MustCompile(`<w:num w:numId="(\d+)">`)
	numMatches := numRe.FindAllStringSubmatch(numXML, -1)
	fmt.Printf("  num elements found: %d\n", len(numMatches))
	for _, m := range numMatches {
		fmt.Printf("  numId=%s\n", m[1])
	}

	// Also show a snippet of the numbering XML around </w:abstractNum>
	idx := strings.Index(numXML, "</w:abstractNum>")
	if idx > 0 {
		// Find the last abstractNum end
		lastIdx := strings.LastIndex(numXML, "</w:abstractNum>")
		if lastIdx > 0 {
			end := lastIdx + 200
			if end > len(numXML) {
				end = len(numXML)
			}
			fmt.Printf("\n=== Tail of numbering.xml (after last abstractNum) ===\n%s\n", numXML[lastIdx:end])
		}
	}
}

func readZipFile(r *zip.ReadCloser, name string) string {
	for _, f := range r.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			var buf strings.Builder
			b := make([]byte, 4096)
			for {
				n, err := rc.Read(b)
				if n > 0 {
					buf.Write(b[:n])
				}
				if err != nil {
					break
				}
			}
			return buf.String()
		}
	}
	return ""
}
