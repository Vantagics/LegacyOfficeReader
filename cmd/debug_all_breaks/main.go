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

	paraRe := regexp.MustCompile(`<w:p[ >].*?</w:p>|<w:p/>`)
	paras := paraRe.FindAllString(docXML, -1)

	fmt.Printf("Total paragraphs: %d\n\n", len(paras))

	for i, p := range paras {
		hasBreak := strings.Contains(p, `w:type="page"`)
		hasPBB := strings.Contains(p, "pageBreakBefore")
		hasSectPr := strings.Contains(p, "w:sectPr")

		if !hasBreak && !hasPBB && !hasSectPr {
			continue
		}

		textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
		texts := textRe.FindAllStringSubmatch(p, -1)
		allText := ""
		for _, t := range texts {
			allText += t[1]
		}
		if len(allText) > 60 {
			allText = allText[:60] + "..."
		}
		if allText == "" {
			allText = "[EMPTY]"
		}

		flags := ""
		if hasBreak {
			flags += " [br:page]"
		}
		if hasPBB {
			flags += " [pageBreakBefore]"
		}
		if hasSectPr {
			flags += " [sectPr]"
		}

		fmt.Printf("[%3d] %s %s\n", i+1, allText, flags)
	}

	// Also check reference
	fmt.Println("\n--- Reference file ---")
	r2, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "ref: %v\n", err)
		return
	}
	defer r2.Close()

	refXML := readZip(r2, "word/document.xml")
	refParas := paraRe.FindAllString(refXML, -1)
	fmt.Printf("Total paragraphs: %d\n\n", len(refParas))

	for i, p := range refParas {
		hasBreak := strings.Contains(p, `w:type="page"`)
		hasPBB := strings.Contains(p, "pageBreakBefore")
		hasSectPr := strings.Contains(p, "w:sectPr")

		if !hasBreak && !hasPBB && !hasSectPr {
			continue
		}

		textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
		texts := textRe.FindAllStringSubmatch(p, -1)
		allText := ""
		for _, t := range texts {
			allText += t[1]
		}
		if len(allText) > 60 {
			allText = allText[:60] + "..."
		}
		if allText == "" {
			allText = "[EMPTY]"
		}

		flags := ""
		if hasBreak {
			flags += " [br:page]"
		}
		if hasPBB {
			flags += " [pageBreakBefore]"
		}
		if hasSectPr {
			flags += " [sectPr]"
		}

		fmt.Printf("[%3d] %s %s\n", i+1, allText, flags)
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
