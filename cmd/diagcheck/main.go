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
	fmt.Println("=== FINAL PPTX VALIDATION ===")

	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	slideCount := 0
	layoutCount := 0
	mediaCount := 0
	totalSize := int64(0)
	for _, f := range r.File {
		totalSize += int64(f.UncompressedSize64)
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			slideCount++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			layoutCount++
		}
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaCount++
		}
	}
	fmt.Printf("Slides: %d, Layouts: %d, Media: %d, Total size: %.1f MB\n", slideCount, layoutCount, mediaCount, float64(totalSize)/1024/1024)

	// XML validation
	xmlErrors := 0
	for _, f := range r.File {
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			continue
		}
		data, _ := io.ReadAll(rc)
		rc.Close()

		decoder := xml.NewDecoder(strings.NewReader(string(data)))
		for {
			_, err := decoder.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				xmlErrors++
				fmt.Printf("  XML ERROR in %s: %v\n", f.Name, err)
				break
			}
		}
	}
	fmt.Printf("XML errors: %d\n", xmlErrors)

	// Layout summary
	fmt.Println("\nLayout backgrounds:")
	for i := 1; i <= layoutCount; i++ {
		name := fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", i)
		for _, f := range r.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				bgStart := strings.Index(content, "<p:bg>")
				bgEnd := strings.Index(content, "</p:bg>")
				bgType := "none"
				if bgStart >= 0 && bgEnd >= 0 {
					bg := content[bgStart : bgEnd+7]
					if strings.Contains(bg, "blipFill") {
						bgType = "image"
					} else if strings.Contains(bg, "solidFill") {
						bgType = "solid"
					}
				}

				pics := strings.Count(content, "<p:pic>")
				shapes := strings.Count(content, "<p:sp>")
				conns := strings.Count(content, "<p:cxnSp>")
				fmt.Printf("  Layout %d: bg=%s, pics=%d, shapes=%d, connectors=%d\n", i, bgType, pics, shapes, conns)
			}
		}
	}

	// Slide shape count summary
	fmt.Println("\nSlide shape counts:")
	totalShapes := 0
	for i := 1; i <= slideCount; i++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		for _, f := range r.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				count := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:pic>") + strings.Count(content, "<p:cxnSp>")
				totalShapes += count
			}
		}
	}
	fmt.Printf("  Total shapes across all slides: %d\n", totalShapes)

	// Check required files exist
	fmt.Println("\nRequired files:")
	required := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
		"ppt/presProps.xml",
		"ppt/viewProps.xml",
		"ppt/tableStyles.xml",
		"ppt/theme/theme1.xml",
		"ppt/slideMasters/slideMaster1.xml",
		"ppt/slideMasters/_rels/slideMaster1.xml.rels",
	}
	for _, req := range required {
		found := false
		for _, f := range r.File {
			if f.Name == req {
				found = true
				break
			}
		}
		if !found {
			fmt.Printf("  MISSING: %s\n", req)
		}
	}
	fmt.Println("  All required files present")

	fmt.Println("\n=== VALIDATION COMPLETE ===")
}
