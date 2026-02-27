package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error:", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/numbering.xml" {
			rc, err := f.Open()
			if err != nil {
				fmt.Println("Error:", err)
				os.Exit(1)
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			// Print with some formatting
			s := string(data)
			s = strings.ReplaceAll(s, "><", ">\n<")
			fmt.Println(s)
			return
		}
	}
	fmt.Println("numbering.xml not found")
}
