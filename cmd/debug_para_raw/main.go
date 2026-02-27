package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
)

func main() {
	path := "testfie/test_new7.docx"
	start := 95
	end := 102

	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "failed to open: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	docXML := readZip(r, "word/document.xml")

	paraRe := regexp.MustCompile(`<w:p[ >].*?</w:p>|<w:p/>`)
	paras := paraRe.FindAllString(docXML, -1)

	for i := start - 1; i < end && i < len(paras); i++ {
		p := paras[i]
		fmt.Printf("\n=== PARA %d (len=%d) ===\n", i+1, len(p))
		if len(p) > 2000 {
			fmt.Printf("%s\n...(truncated, total %d bytes)\n", p[:2000], len(p))
		} else {
			fmt.Printf("%s\n", p)
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
