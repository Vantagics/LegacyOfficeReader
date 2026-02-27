package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/footer1.xml" || f.Name == "word/header1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("=== %s ===\n%s\n\n", f.Name, string(data))
		}
	}

	// Also check document.xml for sectPr references
	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			s := string(data)
			// Find sectPr
			for i := len(s) - 500; i < len(s); i++ {
				if i >= 0 {
					fmt.Print(string(s[i]))
				}
			}
			fmt.Println()
		}
	}
}
