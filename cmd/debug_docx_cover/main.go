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

	// Show paragraphs 1-15 in detail (cover page area)
	for i := 0; i < 15 && i < len(paras); i++ {
		p := paras[i]
		
		textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
		texts := textRe.FindAllStringSubmatch(p, -1)
		allText := ""
		for _, t := range texts {
			allText += t[1]
		}
		
		hasInline := strings.Contains(p, "<wp:inline")
		hasAnchor := strings.Contains(p, "<wp:anchor")
		hasTextBox := strings.Contains(p, "wps:txbx") || strings.Contains(p, "<v:textbox")
		
		fmt.Printf("\n=== PARA %d ===\n", i+1)
		if allText != "" {
			display := allText
			if len(display) > 80 {
				display = display[:80] + "..."
			}
			fmt.Printf("  TEXT: %s\n", display)
		}
		if hasInline {
			fmt.Printf("  HAS INLINE IMAGE\n")
		}
		if hasAnchor {
			fmt.Printf("  HAS ANCHOR IMAGE\n")
		}
		if hasTextBox {
			fmt.Printf("  HAS TEXTBOX\n")
		}
		if allText == "" && !hasInline && !hasAnchor && !hasTextBox {
			fmt.Printf("  ** EMPTY **\n")
		}
		
		snippet := p
		if len(snippet) > 500 {
			snippet = snippet[:500] + "..."
		}
		fmt.Printf("  RAW: %s\n", snippet)
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
