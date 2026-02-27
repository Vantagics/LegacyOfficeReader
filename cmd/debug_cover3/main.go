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

			// Find body content
			bodyStart := strings.Index(content, "<w:body>")
			bodyEnd := strings.Index(content, "</w:body>")
			if bodyStart < 0 || bodyEnd < 0 {
				fmt.Println("No body found")
				return
			}
			body := content[bodyStart+8 : bodyEnd]

			// Show first 5000 chars of body
			if len(body) > 8000 {
				body = body[:8000]
			}
			fmt.Println(body)
		}
	}
}
