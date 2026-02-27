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
	r, err := zip.OpenReader(path)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find the table
			tblStart := strings.Index(content, "<w:tbl>")
			tblEnd := strings.Index(content, "</w:tbl>")
			if tblStart >= 0 && tblEnd >= 0 {
				tbl := content[tblStart : tblEnd+8]
				if len(tbl) > 3000 {
					tbl = tbl[:3000] + "..."
				}
				fmt.Printf("Table:\n%s\n", tbl)
			}
		}
	}
}
