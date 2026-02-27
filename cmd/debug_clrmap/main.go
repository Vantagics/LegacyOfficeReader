package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer zr.Close()

	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") || strings.Contains(f.Name, "_rels") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			continue
		}
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		hasClrMap := strings.Contains(content, "clrMapOvr")
		hasMasterSp := strings.Contains(content, "showMasterSp")
		hasSz0 := strings.Contains(content, `sz="0"`)
		size := len(data)

		if !hasClrMap || !hasMasterSp || hasSz0 {
			fmt.Printf("%s: size=%d clrMapOvr=%v showMasterSp=%v sz0=%v\n", f.Name, size, hasClrMap, hasMasterSp, hasSz0)
			// Show last 200 chars
			if size > 200 {
				fmt.Printf("  TAIL: ...%s\n", content[size-200:])
			}
		}
	}
	fmt.Println("Done")
}
