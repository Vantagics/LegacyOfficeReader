package main

import (
	"archive/zip"
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

	// Check slides 4, 9, 26, 63 (high shape count slides)
	checkSlides := []int{4, 9, 26, 63}
	for _, slideNum := range checkSlides {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Count elements more carefully
				spOpen := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
				picOpen := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
				cxnOpen := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")

				// Count closing tags too
				spClose := strings.Count(content, "</p:sp>")
				picClose := strings.Count(content, "</p:pic>")
				cxnClose := strings.Count(content, "</p:cxnSp>")

				fmt.Printf("Slide %d: sp=%d/%d pic=%d/%d cxn=%d/%d total=%d size=%d bytes\n",
					slideNum, spOpen, spClose, picOpen, picClose, cxnOpen, cxnClose,
					spOpen+picOpen+cxnOpen, len(data))
				break
			}
		}
	}
}
