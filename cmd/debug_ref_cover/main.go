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
	path := "testfie/test.docx"
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

	fmt.Printf("Total paragraphs: %d\n", len(paras))

	for i, p := range paras {
		textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
		texts := textRe.FindAllStringSubmatch(p, -1)
		allText := ""
		for _, t := range texts {
			allText += t[1]
		}

		hasPageBreakBefore := strings.Contains(p, "w:pageBreakBefore")
		hasBreakPage := strings.Contains(p, `w:type="page"`)
		hasInline := strings.Contains(p, "<wp:inline")
		hasAnchor := strings.Contains(p, "<wp:anchor")
		hasTextBox := strings.Contains(p, "wps:txbx") || strings.Contains(p, "<v:textbox")
		hasTbl := strings.Contains(p, "<w:tbl")

		pType := "TEXT"
		if allText == "" && !hasInline && !hasAnchor && !hasTextBox {
			pType = "EMPTY"
		}
		if hasInline || hasAnchor {
			pType = "IMAGE"
		}
		if hasTextBox {
			pType = "TEXTBOX"
		}

		flags := ""
		if hasPageBreakBefore {
			flags += " [pageBreakBefore]"
		}
		if hasBreakPage {
			flags += " [br:page]"
		}
		if hasTbl {
			flags += " [table]"
		}

		display := allText
		if len(display) > 60 {
			display = display[:60] + "..."
		}

		// Only show first 100 paragraphs (cover + some body)
		if i >= 100 {
			break
		}

		fmt.Printf("[%3d] %-10s %s%s\n", i+1, pType, display, flags)
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
