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

	issues := 0

	// Check all XML files are well-formed
	for _, f := range zr.File {
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			fmt.Printf("ISSUE: Cannot open %s: %v\n", f.Name, err)
			issues++
			continue
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			fmt.Printf("ISSUE: Cannot read %s: %v\n", f.Name, err)
			issues++
			continue
		}

		// Check XML well-formedness
		d := xml.NewDecoder(strings.NewReader(string(data)))
		for {
			_, err := d.Token()
			if err != nil {
				if err.Error() == "EOF" {
					break
				}
				fmt.Printf("ISSUE: XML parse error in %s: %v\n", f.Name, err)
				issues++
				break
			}
		}
	}

	// Check all image references in slides point to existing media files
	mediaFiles := make(map[string]bool)
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaFiles[f.Name] = true
		}
	}

	// Check slide rels for broken image references
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/_rels/") && !strings.HasPrefix(f.Name, "ppt/slideLayouts/_rels/") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()

		d := xml.NewDecoder(strings.NewReader(string(data)))
		for {
			tok, err := d.Token()
			if err != nil {
				break
			}
			if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "Relationship" {
				var target, relType string
				for _, attr := range se.Attr {
					switch attr.Name.Local {
					case "Target":
						target = attr.Value
					case "Type":
						relType = attr.Value
					}
				}
				if strings.Contains(relType, "image") {
					// Resolve relative path
					mediaPath := "ppt/media/" + strings.TrimPrefix(target, "../media/")
					if !mediaFiles[mediaPath] {
						fmt.Printf("ISSUE: Broken image ref in %s: %s -> %s\n", f.Name, target, mediaPath)
						issues++
					}
				}
			}
		}
	}

	// Summary
	fmt.Printf("\nFiles checked: %d XML/rels files, %d media files\n", countXML(zr), len(mediaFiles))
	fmt.Printf("Total issues: %d\n", issues)
}

func countXML(zr *zip.ReadCloser) int {
	count := 0
	for _, f := range zr.File {
		if strings.HasSuffix(f.Name, ".xml") || strings.HasSuffix(f.Name, ".rels") {
			count++
		}
	}
	return count
}
