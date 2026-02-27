package main

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide4.xml" {
			rc, _ := f.Open()
			data := make([]byte, f.UncompressedSize64)
			rc.Read(data)
			rc.Close()
			content := string(data)

			// Count different element types
			fmt.Printf("File size: %d bytes\n", len(content))
			fmt.Printf("<p:sp> count: %d\n", strings.Count(content, "<p:sp>"))
			fmt.Printf("<p:pic> count: %d\n", strings.Count(content, "<p:pic>"))
			fmt.Printf("<p:cxnSp> count: %d\n", strings.Count(content, "<p:cxnSp>"))

			// Check for noFill shapes
			fmt.Printf("<a:noFill/> count: %d\n", strings.Count(content, "<a:noFill/>"))

			// Print first 3000 chars to see structure
			if len(content) > 5000 {
				fmt.Println("\n--- First 5000 chars ---")
				fmt.Println(content[:5000])
			} else {
				fmt.Println(content)
			}
			break
		}
	}
}
