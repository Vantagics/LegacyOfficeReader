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

	// Search for specific text
	searches := []string{"文件威胁鉴定器", "其他周边组件", "文档检测", "商业化"}

	paraRe := regexp.MustCompile(`<w:p[ >].*?</w:p>|<w:p/>`)
	paras := paraRe.FindAllString(docXML, -1)

	for _, search := range searches {
		fmt.Printf("\n=== Searching for: %s ===\n", search)
		for i, p := range paras {
			if strings.Contains(p, search) {
				textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)
				texts := textRe.FindAllStringSubmatch(p, -1)
				allText := ""
				for _, t := range texts {
					allText += t[1]
				}
				if len(allText) > 80 {
					allText = allText[:80] + "..."
				}
				hasImg := strings.Contains(p, "<wp:inline") || strings.Contains(p, "<wp:anchor")
				imgFlag := ""
				if hasImg {
					imgFlag = " [IMAGE]"
				}
				fmt.Printf("  [%3d] %s%s\n", i+1, allText, imgFlag)
			}
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
