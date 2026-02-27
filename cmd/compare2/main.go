package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Focus on slide 4 as an example
	slideIdx := 3 // 0-based
	slide := p.GetSlides()[slideIdx]

	fmt.Printf("=== Slide %d PPT texts ===\n", slideIdx+1)
	for i, t := range slide.GetTexts() {
		fmt.Printf("  Text[%d]: %q\n", i, truncate(t, 120))
	}

	fmt.Printf("\n=== Slide %d PPT shapes text ===\n", slideIdx+1)
	for i, sh := range slide.GetShapes() {
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				if strings.TrimSpace(run.Text) != "" {
					fmt.Printf("  Shape[%d]: %q\n", i, truncate(run.Text, 120))
				}
			}
		}
	}

	fmt.Printf("\n=== Slide %d PPTX texts ===\n", slideIdx+1)
	name := fmt.Sprintf("ppt/slides/slide%d.xml", slideIdx+1)
	for _, f := range zr.File {
		if f.Name == name {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			texts := extractTexts(data)
			for i, t := range texts {
				fmt.Printf("  PPTX[%d]: %q\n", i, truncate(t, 120))
			}
		}
	}

	// Now check: which PPT texts are NOT in any shape?
	fmt.Printf("\n=== PPT texts NOT in shapes ===\n")
	for i, t := range slide.GetTexts() {
		t = strings.TrimSpace(t)
		if t == "" {
			continue
		}
		found := false
		for _, sh := range slide.GetShapes() {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if strings.Contains(run.Text, t) || strings.Contains(t, run.Text) {
						found = true
						break
					}
				}
				if found {
					break
				}
			}
			if found {
				break
			}
		}
		if !found {
			// Check if first 15 chars match
			prefix := t
			if len(prefix) > 15 {
				prefix = prefix[:15]
			}
			foundPrefix := false
			for _, sh := range slide.GetShapes() {
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						if strings.Contains(run.Text, prefix) {
							foundPrefix = true
							break
						}
					}
					if foundPrefix {
						break
					}
				}
				if foundPrefix {
					break
				}
			}
			if !foundPrefix {
				fmt.Printf("  Text[%d] NOT in shapes: %q\n", i, truncate(t, 100))
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

func truncate(s string, n int) string {
	s = strings.ReplaceAll(s, "\n", "\\n")
	s = strings.ReplaceAll(s, "\r", "\\r")
	if len(s) > n {
		return s[:n] + "..."
	}
	return s
}
