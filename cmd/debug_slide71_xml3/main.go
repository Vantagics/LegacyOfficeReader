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

	for _, f := range r.File {
		if f.Name == "ppt/slides/slide71.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			content := string(data)
			rc.Close()

			// Find the custGeom shape (Shape 8 = cloud)
			idx := strings.Index(content, `name="Shape 8"`)
			if idx < 0 {
				fmt.Println("Shape 8 not found")
				return
			}
			// Find the enclosing <p:sp>
			spStart := strings.LastIndex(content[:idx], "<p:sp>")
			spEnd := strings.Index(content[spStart:], "</p:sp>")
			if spStart < 0 || spEnd < 0 {
				fmt.Println("Could not find sp boundaries")
				return
			}
			shapeXML := content[spStart : spStart+spEnd+len("</p:sp>")]
			fmt.Printf("Cloud shape (Shape 8) XML:\n%s\n", shapeXML)

			// Also check the line shape (Shape 9 = vertical line)
			fmt.Println("\n---")
			idx2 := strings.Index(content, `name="Connector 9"`)
			if idx2 >= 0 {
				cxnStart := strings.LastIndex(content[:idx2], "<p:cxnSp>")
				cxnEnd := strings.Index(content[cxnStart:], "</p:cxnSp>")
				if cxnStart >= 0 && cxnEnd >= 0 {
					cxnXML := content[cxnStart : cxnStart+cxnEnd+len("</p:cxnSp>")]
					fmt.Printf("Connector 9 XML:\n%s\n", cxnXML)
				}
			}

			// Check images
			fmt.Println("\n---")
			idx3 := strings.Index(content, `name="Image 3"`)
			if idx3 >= 0 {
				picStart := strings.LastIndex(content[:idx3], "<p:pic>")
				picEnd := strings.Index(content[picStart:], "</p:pic>")
				if picStart >= 0 && picEnd >= 0 {
					picXML := content[picStart : picStart+picEnd+len("</p:pic>")]
					fmt.Printf("Image 3 XML:\n%s\n", picXML)
				}
			}
			return
		}
	}
}
