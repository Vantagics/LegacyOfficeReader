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

			// Find "首创使用" (first bullet point in section 4)
			idx := strings.Index(content, "首创使用")
			if idx >= 0 {
				pStart := strings.LastIndex(content[:idx], "<w:p>")
				pEnd := strings.Index(content[idx:], "</w:p>")
				if pStart >= 0 && pEnd >= 0 {
					para := content[pStart : idx+pEnd+len("</w:p>")]
					if len(para) > 600 { para = para[:600] + "..." }
					fmt.Printf("=== Bullet item ===\n%s\n", para)
				}
			}
			
			// Find "检测发现传统" (sub-bullet)
			idx = strings.Index(content, "检测发现传统防护手段无法")
			if idx >= 0 {
				pStart := strings.LastIndex(content[:idx], "<w:p>")
				pEnd := strings.Index(content[idx:], "</w:p>")
				if pStart >= 0 && pEnd >= 0 {
					para := content[pStart : idx+pEnd+len("</w:p>")]
					if len(para) > 600 { para = para[:600] + "..." }
					fmt.Printf("\n=== Sub-bullet item ===\n%s\n", para)
				}
			}
		}
		
		// Check numbering.xml
		if f.Name == "word/numbering.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\n=== numbering.xml ===\n%s\n", string(data))
		}
	}
}
