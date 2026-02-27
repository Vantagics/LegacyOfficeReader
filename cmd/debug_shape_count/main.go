package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

var (
	reSpOpen  = regexp.MustCompile(`<p:sp[ >]`)
	rePicOpen = regexp.MustCompile(`<p:pic[ >]`)
	reCxnOpen = regexp.MustCompile(`<p:cxnSp[ >]`)
)

func countShapes(content string) (sp, pic, cxn int) {
	sp = len(reSpOpen.FindAllString(content, -1))
	pic = len(rePicOpen.FindAllString(content, -1))
	cxn = len(reCxnOpen.FindAllString(content, -1))
	return
}

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	fmt.Printf("%-8s %-6s %-6s %-8s\n", "Slide", "PPT", "PPTX", "Status")
	fmt.Println(strings.Repeat("-", 35))

	mismatchCount := 0
	for i, s := range slides {
		pptCount := len(s.GetShapes())
		slideNum := i + 1

		content := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", slideNum))
		sp, pic, cxn := countShapes(content)
		pptxCount := sp + pic + cxn

		status := "✓"
		if pptCount != pptxCount {
			status = fmt.Sprintf("DIFF %+d", pptxCount-pptCount)
			mismatchCount++
		}

		if pptCount != pptxCount {
			fmt.Printf("Slide %2d  %4d   %4d   %s\n", slideNum, pptCount, pptxCount, status)
		}
	}

	if mismatchCount == 0 {
		fmt.Printf("\nAll %d slides match ✓\n", len(slides))
	} else {
		fmt.Printf("\n%d/%d slides have mismatches\n", mismatchCount, len(slides))
	}
}

func readZipFile(zr *zip.ReadCloser, name string) string {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			return string(data)
		}
	}
	return ""
}
