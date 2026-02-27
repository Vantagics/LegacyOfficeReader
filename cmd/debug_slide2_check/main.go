package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
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
		if f.Name == "ppt/slides/slide2.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()

			// Extract all text content
			d := xml.NewDecoder(strings.NewReader(string(data)))
			inT := false
			var texts []string
			for {
				tok, err := d.Token()
				if err != nil {
					break
				}
				switch t := tok.(type) {
				case xml.StartElement:
					if t.Name.Local == "t" {
						inT = true
					}
				case xml.EndElement:
					if t.Name.Local == "t" {
						inT = false
					}
				case xml.CharData:
					if inT {
						texts = append(texts, string(t))
					}
				}
			}

			fmt.Printf("Slide 2 texts (%d):\n", len(texts))
			for i, t := range texts {
				fmt.Printf("  [%d] %q\n", i, t)
			}
			break
		}
	}
}
