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
		if zf.Name == "ppt/slideLayouts/slideLayout4.xml" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Show all shapes
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

				nameIdx := strings.Index(content[absIdx:], `name="`)
				name := ""
				if nameIdx >= 0 && nameIdx < 200 {
					nameEnd := strings.Index(content[absIdx+nameIdx+6:], `"`)
					if nameEnd >= 0 {
						name = content[absIdx+nameIdx+6 : absIdx+nameIdx+6+nameEnd]
					}
				}

				offIdx := strings.Index(content[absIdx:], `<a:off `)
				pos := ""
				if offIdx >= 0 && offIdx < 500 {
					offEnd := strings.Index(content[absIdx+offIdx:], "/>")
					if offEnd >= 0 {
						pos = content[absIdx+offIdx : absIdx+offIdx+offEnd+2]
					}
				}

				embedIdx := strings.Index(content[absIdx:absIdx+min(1000, len(content)-absIdx)], `r:embed="`)
				embed := ""
				if embedIdx >= 0 {
					embedEnd := strings.Index(content[absIdx+embedIdx+9:], `"`)
					if embedEnd >= 0 {
						embed = content[absIdx+embedIdx+9 : absIdx+embedIdx+9+embedEnd]
					}
				}

				extra := ""
				if embed != "" {
					extra = fmt.Sprintf(" [IMG:%s]", embed)
				}

				fmt.Printf("z=%2d type=%-5s name=%-30s %s%s\n", order, elemType, name, pos, extra)
				idx = absIdx + 1
			}

			// Check showMasterSp
			if strings.Contains(content, `showMasterSp="0"`) {
				fmt.Println("\nshowMasterSp=0")
			} else if strings.Contains(content, `showMasterSp="1"`) {
				fmt.Println("\nshowMasterSp=1")
			}

			// Check rels
		}

		if zf.Name == "ppt/slideLayouts/_rels/slideLayout4.xml.rels" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\nLayout4 rels:\n%s\n", string(data))
		}
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
