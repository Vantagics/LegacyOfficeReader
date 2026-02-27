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

			// Show first 5000 chars of the body
			bodyStart := strings.Index(content, "<w:body>")
			if bodyStart < 0 {
				fmt.Println("No <w:body> found")
				return
			}
			body := content[bodyStart:]
			if len(body) > 6000 {
				body = body[:6000]
			}
			fmt.Println(body)
		}
	}
}
