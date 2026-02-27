package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	// Tiger watermark slides: 8,9,11,12,13,14,15,17,18,19,20,22,23,24,25,26,27,29,30,31,32,33,34,36,37,38,39,40,42,43,44,45,46,47,49,50,51,52,53,54,55,56,57,59,60,61,62,63,64,65,66,67,68,69,70,71
	tigerSlides := []int{8, 9, 11, 12, 13, 14, 15, 17, 18, 19, 20, 22, 23, 24, 25, 26, 27, 29, 30, 31, 32, 33, 34, 36, 37, 38, 39, 40, 42, 43, 44, 45, 46, 47, 49, 50, 51, 52, 53, 54, 55, 56, 57, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71}

	missing := 0
	present := 0
	for _, slideNum := range tigerSlides {
		slideName := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		for _, zf := range r.File {
			if zf.Name != slideName {
				continue
			}
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			if strings.Contains(content, "rImg14") {
				present++
			} else {
				missing++
				fmt.Printf("Slide %d: MISSING watermark (rImg14)\n", slideNum)
			}
		}
	}
	fmt.Printf("\nTotal: %d present, %d missing\n", present, missing)
}
