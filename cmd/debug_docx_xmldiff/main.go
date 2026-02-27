package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"sort"
	"strings"
)

func readZipFile(zipPath, innerPath string) ([]byte, error) {
	r, err := zip.OpenReader(zipPath)
	if err != nil {
		return nil, err
	}
	defer r.Close()
	for _, f := range r.File {
		if f.Name == innerPath {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("file %q not found in %s", innerPath, zipPath)
}

func listZipFiles(zipPath string) ([]string, error) {
	r, err := zip.OpenReader(zipPath)
	if err != nil {
		return nil, err
	}
	defer r.Close()
	var names []string
	for _, f := range r.File {
		names = append(names, f.Name)
	}
	sort.Strings(names)
	return names, nil
}

func main() {
	ref := "testfie/test.docx"
	our := "testfie/test_new8.docx"

	// List all files in both
	refFiles, _ := listZipFiles(ref)
	ourFiles, _ := listZipFiles(our)

	refSet := map[string]bool{}
	for _, f := range refFiles {
		refSet[f] = true
	}
	ourSet := map[string]bool{}
	for _, f := range ourFiles {
		ourSet[f] = true
	}

	fmt.Println("=== Files only in reference ===")
	for _, f := range refFiles {
		if !ourSet[f] {
			fmt.Println("  ", f)
		}
	}
	fmt.Println("=== Files only in our output ===")
	for _, f := range ourFiles {
		if !refSet[f] {
			fmt.Println("  ", f)
		}
	}

	// Compare common files
	fmt.Println("\n=== File size comparison ===")
	for _, f := range refFiles {
		if !ourSet[f] {
			continue
		}
		refData, _ := readZipFile(ref, f)
		ourData, _ := readZipFile(our, f)
		if bytes.Equal(refData, ourData) {
			fmt.Printf("  SAME  %s (%d bytes)\n", f, len(refData))
		} else {
			fmt.Printf("  DIFF  %s (ref=%d, our=%d)\n", f, len(refData), len(ourData))
		}
	}

	// Detailed XML diff for document.xml
	fmt.Println("\n=== document.xml detailed diff ===")
	refDoc, _ := readZipFile(ref, "word/document.xml")
	ourDoc, _ := readZipFile(our, "word/document.xml")

	if bytes.Equal(refDoc, ourDoc) {
		fmt.Println("document.xml is IDENTICAL")
		return
	}

	// Split by paragraph tags and compare
	refParas := splitXMLParas(string(refDoc))
	ourParas := splitXMLParas(string(ourDoc))

	fmt.Printf("Reference paragraphs: %d\n", len(refParas))
	fmt.Printf("Our paragraphs: %d\n", len(ourParas))

	maxParas := len(refParas)
	if len(ourParas) > maxParas {
		maxParas = len(ourParas)
	}

	diffCount := 0
	for i := 0; i < maxParas; i++ {
		var rp, op string
		if i < len(refParas) {
			rp = refParas[i]
		}
		if i < len(ourParas) {
			op = ourParas[i]
		}
		if rp != op {
			diffCount++
			if diffCount <= 20 {
				fmt.Printf("\n--- Para %d DIFF ---\n", i)
				if len(rp) > 500 {
					fmt.Printf("REF[%d]: %s...\n", len(rp), rp[:500])
				} else {
					fmt.Printf("REF[%d]: %s\n", len(rp), rp)
				}
				if len(op) > 500 {
					fmt.Printf("OUR[%d]: %s...\n", len(op), op[:500])
				} else {
					fmt.Printf("OUR[%d]: %s\n", len(op), op)
				}
			}
		}
	}
	fmt.Printf("\nTotal paragraph diffs: %d\n", diffCount)
}

func splitXMLParas(xml string) []string {
	// Split on <w:p or <w:tbl boundaries
	var result []string
	rest := xml
	for {
		// Find next <w:p or <w:tbl
		pIdx := strings.Index(rest, "<w:p ")
		p2Idx := strings.Index(rest, "<w:p>")
		tIdx := strings.Index(rest, "<w:tbl")

		minIdx := -1
		if pIdx >= 0 {
			minIdx = pIdx
		}
		if p2Idx >= 0 && (minIdx < 0 || p2Idx < minIdx) {
			minIdx = p2Idx
		}
		if tIdx >= 0 && (minIdx < 0 || tIdx < minIdx) {
			minIdx = tIdx
		}

		if minIdx < 0 {
			break
		}

		// Find the end of this element
		if minIdx == tIdx {
			endTag := "</w:tbl>"
			endIdx := strings.Index(rest[minIdx:], endTag)
			if endIdx < 0 {
				result = append(result, rest[minIdx:])
				break
			}
			result = append(result, rest[minIdx:minIdx+endIdx+len(endTag)])
			rest = rest[minIdx+endIdx+len(endTag):]
		} else {
			// It's a <w:p> or <w:p ...>
			endTag := "</w:p>"
			endIdx := strings.Index(rest[minIdx:], endTag)
			if endIdx < 0 {
				result = append(result, rest[minIdx:])
				break
			}
			result = append(result, rest[minIdx:minIdx+endIdx+len(endTag)])
			rest = rest[minIdx+endIdx+len(endTag):]
		}
	}
	return result
}
