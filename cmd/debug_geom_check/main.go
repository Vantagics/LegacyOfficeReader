package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	f, _ := zip.OpenReader("testfie/test.pptx")
	defer f.Close()

	for _, zf := range f.File {
		if !strings.HasPrefix(zf.Name, "ppt/slides/slide") || !strings.HasSuffix(zf.Name, ".xml") {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		xml := string(data)

		custCount := strings.Count(xml, "custGeom")
		prstCount := strings.Count(xml, "prstGeom")
		if custCount > 0 {
			fmt.Printf("%s: custGeom=%d prstGeom=%d\n", zf.Name, custCount, prstCount)
		}
	}
}
