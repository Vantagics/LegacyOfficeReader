package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
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

	slideCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideCount++
		}
	}
	fmt.Printf("PPTX slide count: %d\n", slideCount)

	// Print text from each slide
	for i := 1; i <= slideCount; i++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data := make([]byte, f.UncompressedSize64)
				rc.Read(data)
				rc.Close()
				texts := extractTexts(data)
				preview := ""
				for _, t := range texts {
					if len(t) > 0 {
						preview = t
						break
					}
				}
				if len(preview) > 60 {
					preview = preview[:60] + "..."
				}
				fmt.Printf("Slide %d: %d texts, preview=%q\n", i, len(texts), preview)
				break
			}
		}
	}
}

func extractTexts(xmlData []byte) []string {
	type AText struct {
		Value string `xml:",chardata"`
	}
	decoder := xml.NewDecoder(strings.NewReader(string(xmlData)))
	var texts []string
	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		if se, ok := tok.(xml.StartElement); ok && se.Name.Local == "t" {
			var at AText
			if err := decoder.DecodeElement(&at, &se); err == nil && at.Value != "" {
				texts = append(texts, at.Value)
			}
		}
	}
	return texts
}
