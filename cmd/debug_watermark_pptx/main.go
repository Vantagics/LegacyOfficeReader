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

	// Check slide8 for watermark image references
	for _, zf := range r.File {
		if zf.Name == "ppt/slides/slide8.xml" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find all pic elements
			parts := strings.Split(content, "<p:pic>")
			fmt.Printf("Slide 8: %d pic elements\n", len(parts)-1)
			for i, part := range parts {
				if i == 0 {
					continue
				}
				// Get name
				nameIdx := strings.Index(part, `name="`)
				name := ""
				if nameIdx >= 0 {
					nameEnd := strings.Index(part[nameIdx+6:], `"`)
					if nameEnd >= 0 {
						name = part[nameIdx+6 : nameIdx+6+nameEnd]
					}
				}
				// Get embed ref
				embedIdx := strings.Index(part, `r:embed="`)
				embed := ""
				if embedIdx >= 0 {
					embedEnd := strings.Index(part[embedIdx+9:], `"`)
					if embedEnd >= 0 {
						embed = part[embedIdx+9 : embedIdx+9+embedEnd]
					}
				}
				// Get position
				offIdx := strings.Index(part, `<a:off `)
				pos := ""
				if offIdx >= 0 {
					offEnd := strings.Index(part[offIdx:], "/>")
					if offEnd >= 0 {
						pos = part[offIdx : offIdx+offEnd+2]
					}
				}
				extIdx := strings.Index(part, `<a:ext `)
				size := ""
				if extIdx >= 0 {
					extEnd := strings.Index(part[extIdx:], "/>")
					if extEnd >= 0 {
						size = part[extIdx : extIdx+extEnd+2]
					}
				}
				fmt.Printf("  pic[%d]: name=%s embed=%s\n    %s\n    %s\n", i, name, embed, pos, size)
			}
		}

		// Check slide8 rels
		if zf.Name == "ppt/slides/_rels/slide8.xml.rels" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			fmt.Printf("\nSlide 8 rels:\n%s\n", string(data))
		}
	}

	// Check layout files for watermark
	for _, zf := range r.File {
		if strings.HasPrefix(zf.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(zf.Name, ".xml") {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if strings.Contains(content, "pic") {
				picCount := strings.Count(content, "<p:pic>")
				fmt.Printf("\n%s: %d pic elements\n", zf.Name, picCount)
				// Show image details
				parts := strings.Split(content, "<p:pic>")
				for i, part := range parts {
					if i == 0 {
						continue
					}
					embedIdx := strings.Index(part, `r:embed="`)
					embed := ""
					if embedIdx >= 0 {
						embedEnd := strings.Index(part[embedIdx+9:], `"`)
						if embedEnd >= 0 {
							embed = part[embedIdx+9 : embedIdx+9+embedEnd]
						}
					}
					offIdx := strings.Index(part, `<a:off `)
					pos := ""
					if offIdx >= 0 {
						offEnd := strings.Index(part[offIdx:], "/>")
						if offEnd >= 0 {
							pos = part[offIdx : offIdx+offEnd+2]
						}
					}
					extIdx := strings.Index(part, `<a:ext `)
					size := ""
					if extIdx >= 0 {
						extEnd := strings.Index(part[extIdx:], "/>")
						if extEnd >= 0 {
							size = part[extIdx : extIdx+extEnd+2]
						}
					}
					fmt.Printf("  pic[%d]: embed=%s\n    %s\n    %s\n", i, embed, pos, size)
				}
			}
		}
	}

	// Check what images are in the pptx
	fmt.Println("\n\nAll images in pptx:")
	for _, zf := range r.File {
		if strings.HasPrefix(zf.Name, "ppt/media/") {
			fmt.Printf("  %s (size=%d)\n", zf.Name, zf.UncompressedSize64)
		}
	}
}
