package main

import (
	"archive/zip"
	"encoding/xml"
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

	for i := 1; i <= 7; i++ {
		fname := fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", i)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				fmt.Printf("=== Layout %d (%d bytes) ===\n", i, len(data))

				// Check background
				if strings.Contains(content, "<p:bg>") {
					if strings.Contains(content, "blipFill") {
						// Find the rImg reference
						idx := strings.Index(content, "r:embed=\"")
						if idx >= 0 {
							end := strings.Index(content[idx+9:], "\"")
							if end >= 0 {
								fmt.Printf("  Background: blipFill ref=%s\n", content[idx+9:idx+9+end])
							}
						}
					} else if strings.Contains(content, "solidFill") {
						idx := strings.Index(content, "srgbClr val=\"")
						if idx >= 0 {
							fmt.Printf("  Background: solidFill color=%s\n", content[idx+13:idx+19])
						}
					}
				} else {
					fmt.Println("  Background: none")
				}

				// Count shapes
				spCount := strings.Count(content, "</p:sp>")
				picCount := strings.Count(content, "</p:pic>")
				cxnCount := strings.Count(content, "</p:cxnSp>")
				fmt.Printf("  Shapes: sp=%d pic=%d cxn=%d\n", spCount, picCount, cxnCount)

				// Extract image refs from pics
				d := xml.NewDecoder(strings.NewReader(content))
				for {
					tok, err := d.Token()
					if err != nil {
						break
					}
					if se, ok := tok.(xml.StartElement); ok {
						if se.Name.Local == "blip" {
							for _, attr := range se.Attr {
								if attr.Name.Local == "embed" {
									fmt.Printf("  Image ref: %s\n", attr.Value)
								}
							}
						}
					}
				}

				// Check rels
				relsFname := fmt.Sprintf("ppt/slideLayouts/_rels/slideLayout%d.xml.rels", i)
				for _, rf := range zr.File {
					if rf.Name == relsFname {
						rc2, _ := rf.Open()
						rdata, _ := io.ReadAll(rc2)
						rc2.Close()
						rcontent := string(rdata)
						imgCount := strings.Count(rcontent, "relationships/image")
						fmt.Printf("  Rels: %d image refs\n", imgCount)
						// Show image targets
						d2 := xml.NewDecoder(strings.NewReader(rcontent))
						for {
							tok, err := d2.Token()
							if err != nil {
								break
							}
							if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "Relationship" {
								var id, target, relType string
								for _, attr := range se.Attr {
									switch attr.Name.Local {
									case "Id":
										id = attr.Value
									case "Target":
										target = attr.Value
									case "Type":
										relType = attr.Value
									}
								}
								if strings.Contains(relType, "image") {
									fmt.Printf("    %s -> %s\n", id, target)
								}
							}
						}
						break
					}
				}
				fmt.Println()
				break
			}
		}
	}
}
