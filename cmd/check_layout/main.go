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

	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()

			content := string(data)
			hasBg := strings.Contains(content, "<p:bg>")
			hasBlip := strings.Contains(content, "blipFill")
			spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
			picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")

			// Extract text content
			var texts []string
			d := xml.NewDecoder(strings.NewReader(content))
			inT := false
			for {
				tok, err := d.Token()
				if err != nil {
					break
				}
				switch t := tok.(type) {
				case xml.StartElement:
					if t.Name.Local == "t" {
						inT = true
					}
				case xml.EndElement:
					if t.Name.Local == "t" {
						inT = false
					}
				case xml.CharData:
					if inT {
						s := strings.TrimSpace(string(t))
						if s != "" {
							texts = append(texts, s)
						}
					}
				}
			}

			fmt.Printf("%s: bg=%v blip=%v shapes=%d pics=%d texts=%d\n", f.Name, hasBg, hasBlip, spCount, picCount, len(texts))
			for _, t := range texts {
				if len(t) > 60 {
					t = t[:60] + "..."
				}
				fmt.Printf("  text: %q\n", t)
			}
		}
	}

	// Check which layout each slide references
	fmt.Println("\nSlide -> Layout mapping:")
	for i := 1; i <= 10; i++ {
		relsName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", i)
		for _, f := range zr.File {
			if f.Name == relsName {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				// Find layout reference
				idx := strings.Index(content, "slideLayout")
				if idx >= 0 {
					end := strings.Index(content[idx:], `"`)
					if end > 0 {
						fmt.Printf("  Slide %d -> %s\n", i, content[idx:idx+end])
					}
				}
			}
		}
	}
}
