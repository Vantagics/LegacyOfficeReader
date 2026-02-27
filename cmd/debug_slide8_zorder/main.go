package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
	"strconv"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, zf := range r.File {
		if zf.Name != "ppt/slides/slide8.xml" {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Extract all shape elements with their positions
		// Find <p:pic> and <p:sp> elements
		reOff := regexp.MustCompile(`<a:off x="(\d+)" y="(\d+)"`)
		reExt := regexp.MustCompile(`<a:ext cx="(\d+)" cy="(\d+)"`)
		reBlip := regexp.MustCompile(`r:embed="(rImg\d+)"`)
		reName := regexp.MustCompile(`name="([^"]*)"`)

		// Split by shape elements
		elements := []string{"<p:pic>", "<p:sp>", "<p:cxnSp>"}
		for _, elem := range elements {
			idx := 0
			for {
				pos := strings.Index(content[idx:], elem)
				if pos < 0 {
					break
				}
				start := idx + pos
				// Find end of this element
				endTag := strings.Replace(elem, "<", "</", 1)
				endPos := strings.Index(content[start:], endTag)
				if endPos < 0 {
					break
				}
				fragment := content[start : start+endPos+len(endTag)]

				// Extract info
				name := ""
				if m := reName.FindStringSubmatch(fragment); m != nil {
					name = m[1]
				}
				x, y, cx, cy := 0, 0, 0, 0
				if m := reOff.FindStringSubmatch(fragment); m != nil {
					x, _ = strconv.Atoi(m[1])
					y, _ = strconv.Atoi(m[2])
				}
				if m := reExt.FindStringSubmatch(fragment); m != nil {
					cx, _ = strconv.Atoi(m[1])
					cy, _ = strconv.Atoi(m[2])
				}
				blip := ""
				if m := reBlip.FindStringSubmatch(fragment); m != nil {
					blip = m[1]
				}
				hasCustGeom := strings.Contains(fragment, "<a:custGeom>")
				fill := ""
				if strings.Contains(fragment, "<a:noFill/>") {
					fill = "noFill"
				} else if m := regexp.MustCompile(`<a:solidFill><a:srgbClr val="([^"]+)"`).FindStringSubmatch(fragment); m != nil {
					fill = m[1]
				}

				fmt.Printf("%s name=%q pos=(%d,%d) sz=(%d,%d) blip=%s fill=%s custGeom=%v\n",
					elem, name, x, y, cx, cy, blip, fill, hasCustGeom)

				idx = start + endPos + len(endTag)
			}
		}
	}
}
