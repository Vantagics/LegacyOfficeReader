package main

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error opening: %v\n", err)
				return
			}
			buf := make([]byte, 500)
			n, _ := rc.Read(buf)
			rc.Close()
			
			content := string(buf[:n])
			// Find first eastAsia attribute
			idx := strings.Index(content, "eastAsia")
			if idx >= 0 {
				start := idx
				if start > 20 {
					start = idx - 20
				}
				end := idx + 60
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("Around eastAsia: %q\n", content[start:end])
				// Print raw bytes
				fmt.Printf("Raw bytes: %x\n", []byte(content[start:end]))
			}
		}
	}

	// Also check: write a simple test
	fmt.Printf("\nTest: 宋体 in UTF-8 bytes: %x\n", []byte("宋体"))
	fmt.Printf("Test: 瀹嬩綋 in UTF-8 bytes: %x\n", []byte("瀹嬩綋"))
}
