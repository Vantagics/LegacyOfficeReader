package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	path := "testfie/test.docx"
	if len(os.Args) > 1 {
		path = os.Args[1]
	}

	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	fmt.Println("=== FILES IN DOCX ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	// Print key XML files
	for _, name := range []string{
		"word/document.xml",
		"word/_rels/document.xml.rels",
		"word/styles.xml",
		"word/header1.xml",
		"word/footer1.xml",
		"[Content_Types].xml",
	} {
		for _, f := range r.File {
			if f.Name == name {
				rc, err := f.Open()
				if err != nil {
					continue
				}
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				// Truncate very long files
				if len(content) > 3000 && !strings.Contains(name, "header") && !strings.Contains(name, "footer") && !strings.Contains(name, "rels") && !strings.Contains(name, "Content_Types") {
					content = content[:1500] + "\n...[TRUNCATED]...\n" + content[len(content)-1500:]
				}
				fmt.Printf("\n=== %s ===\n%s\n", name, content)
			}
		}
	}
}
