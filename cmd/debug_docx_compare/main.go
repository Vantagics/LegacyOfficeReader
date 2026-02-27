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

	fmt.Println("=== DOCX ZIP contents ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Read and print key XML files
	for _, name := range []string{"word/document.xml", "word/styles.xml", "[Content_Types].xml", "word/_rels/document.xml.rels", "word/numbering.xml", "word/settings.xml"} {
		for _, f := range r.File {
			if f.Name == name {
				rc, err := f.Open()
				if err != nil {
					continue
				}
				buf := make([]byte, f.UncompressedSize64)
				rc.Read(buf)
				rc.Close()
				content := string(buf)
				// Truncate very long files
				if len(content) > 5000 {
					content = content[:5000] + "\n... [TRUNCATED]"
				}
				fmt.Printf("\n=== %s ===\n%s\n", name, content)
			}
		}
	}

	// Print headers and footers
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "word/header") || strings.HasPrefix(f.Name, "word/footer") {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			buf := make([]byte, f.UncompressedSize64)
			rc.Read(buf)
			rc.Close()
			fmt.Printf("\n=== %s ===\n%s\n", f.Name, string(buf))
		}
	}
}
