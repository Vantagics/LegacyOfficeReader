package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Search for key text
			keywords := []string{"版权声明", "修订记录", "目  录", "目 ", "状态：C", "地址："}
			for _, kw := range keywords {
				idx := strings.Index(content, kw)
				if idx >= 0 {
					// Find the enclosing <w:p>
					pStart := strings.LastIndex(content[:idx], "<w:p>")
					pEnd := strings.Index(content[idx:], "</w:p>")
					if pStart >= 0 && pEnd >= 0 {
						para := content[pStart : idx+pEnd+len("</w:p>")]
						if len(para) > 500 {
							para = para[:500] + "..."
						}
						fmt.Printf("\n=== '%s' ===\n%s\n", kw, para)
					}
				}
			}
		}
	}
}
