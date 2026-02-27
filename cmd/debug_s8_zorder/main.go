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

	for _, zf := range r.File {
		if zf.Name != "ppt/slides/slide8.xml" {
			continue
		}
		rc, _ := zf.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Split by shape/pic elements to see z-order
		idx := 0
		order := 0
		for idx < len(content) {
			spIdx := strings.Index(content[idx:], "<p:sp>")
			picIdx := strings.Index(content[idx:], "<p:pic>")
			cxnIdx := strings.Index(content[idx:], "<p:cxnSp>")

			minIdx := -1
			elemType := ""
			if spIdx >= 0 {
				minIdx = spIdx
				elemType = "sp"
			}
			if picIdx >= 0 && (minIdx < 0 || picIdx < minIdx) {
				minIdx = picIdx
				elemType = "pic"
			}
			if cxnIdx >= 0 && (minIdx < 0 || cxnIdx < minIdx) {
				minIdx = cxnIdx
				elemType = "cxnSp"
			}

			if minIdx < 0 {
				break
			}

			absIdx := idx + minIdx
			order++

			// Get name
			nameIdx := strings.Index(content[absIdx:], `name="`)
			name := ""
			if nameIdx >= 0 && nameIdx < 200 {
				nameEnd := strings.Index(content[absIdx+nameIdx+6:], `"`)
				if nameEnd >= 0 {
					name = content[absIdx+nameIdx+6 : absIdx+nameIdx+6+nameEnd]
				}
			}

			// Get position
			offIdx := strings.Index(content[absIdx:], `<a:off `)
			pos := ""
			if offIdx >= 0 && offIdx < 500 {
				offEnd := strings.Index(content[absIdx+offIdx:], "/>")
				if offEnd >= 0 {
					pos = content[absIdx+offIdx : absIdx+offIdx+offEnd+2]
				}
			}

			// Check for custGeom
			hasCustGeom := false
			nextSp := strings.Index(content[absIdx+1:], "<p:sp>")
			nextPic := strings.Index(content[absIdx+1:], "<p:pic>")
			boundary := len(content) - absIdx - 1
			if nextSp >= 0 && nextSp < boundary {
				boundary = nextSp
			}
			if nextPic >= 0 && nextPic < boundary {
				boundary = nextPic
			}
			segment := content[absIdx : absIdx+1+boundary]
			if strings.Contains(segment, "custGeom") {
				hasCustGeom = true
			}

			// Check for embed (image)
			embedIdx := strings.Index(content[absIdx:absIdx+1+boundary], `r:embed="`)
			embed := ""
			if embedIdx >= 0 {
				embedEnd := strings.Index(content[absIdx+embedIdx+9:], `"`)
				if embedEnd >= 0 {
					embed = content[absIdx+embedIdx+9 : absIdx+embedIdx+9+embedEnd]
				}
			}

			extra := ""
			if hasCustGeom {
				extra += " [FREEFORM]"
			}
			if embed != "" {
				extra += fmt.Sprintf(" [IMG:%s]", embed)
			}

			fmt.Printf("z=%2d type=%-5s name=%-20s %s%s\n", order, elemType, name, pos, extra)

			idx = absIdx + 1
		}
	}
}
