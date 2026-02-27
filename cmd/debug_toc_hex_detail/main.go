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
		if zf.Name == "ppt/slides/slide2.xml" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			xml := string(data)

			idx := 0
			spNum := 0
			for {
				start := strings.Index(xml[idx:], "<p:sp>")
				if start < 0 {
					break
				}
				start += idx
				end := strings.Index(xml[start:], "</p:sp>")
				if end < 0 {
					break
				}
				end += start + len("</p:sp>")
				snippet := xml[start:end]
				spNum++

				// Only show number shapes
				for _, num := range []string{">1<", ">2<", ">3<", ">4<"} {
					if strings.Contains(snippet, num+"/a:t>") {
						fmt.Printf("=== Shape %d (number) ===\n", spNum)
						spPrStart := strings.Index(snippet, "<p:spPr>")
						spPrEnd := strings.Index(snippet, "</p:spPr>")
						if spPrStart >= 0 && spPrEnd >= 0 {
							spPr := snippet[spPrStart : spPrEnd+len("</p:spPr>")]
							fmt.Println(spPr)
						}
						fmt.Println()
					}
				}

				idx = end
			}
		}
	}
}
