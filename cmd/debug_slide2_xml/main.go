package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	// Check slide 2 XML (CONTENTS slide)
	for _, f := range r.File {
		if f.Name == "ppt/slides/slide2.xml" {
			rc, _ := f.Open()
			var buf bytes.Buffer
			buf.ReadFrom(rc)
			rc.Close()
			content := buf.String()
			fmt.Printf("Slide 2 XML (%d bytes):\n", len(content))
			if len(content) > 5000 {
				fmt.Println(content[:5000])
				fmt.Printf("... (%d more bytes)\n", len(content)-5000)
			} else {
				fmt.Println(content)
			}
		}
	}
}
