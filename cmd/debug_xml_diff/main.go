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
	if len(os.Args) < 2 {
		fmt.Println("usage: debug_xml_diff <file.docx> [start] [end]")
		os.Exit(1)
	}
	path := os.Args[1]

	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "failed to open: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	docXML := readZip(r, "word/document.xml")

	// Split into paragraphs and tables
	paraRe := regexp.MustCompile(`<w:p[ >].*?</w:p>|<w:p/>|<w:tbl>.*?</w:tbl>`)
	paras := paraRe.FindAllString(docXML, -1)

	start := 0
	end := len(paras)
	if len(os.Args) > 2 {
		fmt.Sscanf(os.Args[2], "%d", &start)
		start-- // 1-based to 0-based
	}
	if len(os.Args) > 3 {
		fmt.Sscanf(os.Args[3], "%d", &end)
	}

	for i := start; i < end && i < len(paras); i++ {
		p := paras[i]
		// Extract text
		textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
		texts := textRe.FindAllStringSubmatch(p, -1)
		allText := ""
		for _, t := range texts {
			allText += t[1]
		}

		// Extract pPr (paragraph properties)
		pprRe := regexp.MustCompile(`<w:pPr>(.*?)</w:pPr>`)
		pprMatch := pprRe.FindStringSubmatch(p)
		ppr := ""
		if pprMatch != nil {
			ppr = pprMatch[1]
		}

		display := allText
		if len(display) > 50 {
			display = display[:50] + "..."
		}
		if display == "" {
			if strings.Contains(p, "<wp:inline") || strings.Contains(p, "<wp:anchor") {
				display = "[IMAGE]"
			} else if strings.Contains(p, "wps:txbx") || strings.Contains(p, "<v:textbox") {
				display = "[TEXTBOX]"
			} else {
				display = "[EMPTY]"
			}
		}

		fmt.Printf("\n--- PARA %d ---\n", i+1)
		fmt.Printf("  TEXT: %s\n", display)
		if ppr != "" {
			fmt.Printf("  PPR:  %s\n", ppr)
		}
		if strings.Contains(p, `w:type="page"`) {
			fmt.Printf("  ** HAS PAGE BREAK **\n")
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
