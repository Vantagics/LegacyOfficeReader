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

	// Show all paragraphs in detail
	for i := 0; i < len(paras); i++ {
		p := paras[i]
		
		// Extract text
		textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
		texts := textRe.FindAllStringSubmatch(p, -1)
		allText := ""
		for _, t := range texts {
			allText += t[1]
		}
		
		// Check for images
		hasInline := strings.Contains(p, "<wp:inline")
		hasAnchor := strings.Contains(p, "<wp:anchor")
		hasDrawing := strings.Contains(p, "<w:drawing")
		
		// Check for page break
		hasPageBreak := strings.Contains(p, `w:type="page"`)
		
		// Get image dimensions if present
		cxRe := regexp.MustCompile(`cx="(\d+)"`)
		cyRe := regexp.MustCompile(`cy="(\d+)"`)
		cxMatches := cxRe.FindAllStringSubmatch(p, -1)
		cyMatches := cyRe.FindAllStringSubmatch(p, -1)
		
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
		if hasDrawing {
			fmt.Printf("  HAS DRAWING\n")
		}
		if hasPageBreak {
			fmt.Printf("  HAS PAGE BREAK\n")
		}
		if len(cxMatches) > 0 || len(cyMatches) > 0 {
			for j, cx := range cxMatches {
				cy := ""
				if j < len(cyMatches) {
					cy = cyMatches[j][1]
				}
				fmt.Printf("  IMG SIZE: cx=%s cy=%s (%.1fcm x %.1fcm)\n", 
					cx[1], cy, 
					float64(atoi(cx[1]))/914400.0*2.54,
					float64(atoi(cy))/914400.0*2.54)
			}
		}
		if allText == "" && !hasInline && !hasAnchor && !hasPageBreak {
			fmt.Printf("  ** EMPTY **\n")
		}
		
		// Show raw XML snippet (first 300 chars)
		snippet := p
		if len(snippet) > 400 {
			snippet = snippet[:400] + "..."
		}
		fmt.Printf("  RAW: %s\n", snippet)
	}
}

func atoi(s string) int {
	n := 0
	for _, c := range s {
		if c >= '0' && c <= '9' {
			n = n*10 + int(c-'0')
		}
	}
	return n
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
