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
	// Parse PPT
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "PPT error: %v\n", err)
		os.Exit(1)
	}
	slides := p.GetSlides()

	// Parse PPTX text
	zf, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "PPTX error: %v\n", err)
		os.Exit(1)
	}
	defer zf.Close()

	pptxTexts := make(map[int]string)
	for _, file := range zf.File {
		if !strings.HasPrefix(file.Name, "ppt/slides/slide") || !strings.HasSuffix(file.Name, ".xml") {
			continue
		}
		numStr := file.Name[len("ppt/slides/slide") : len(file.Name)-4]
		var num int
		fmt.Sscanf(numStr, "%d", &num)

		rc, _ := file.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()

		// Extract text from XML
		text := extractXMLText(string(data))
		pptxTexts[num] = text
	}

	// Compare
	mismatches := 0
	for i, s := range slides {
		shapes := s.GetShapes()
		var pptText strings.Builder
		for _, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					pptText.WriteString(run.Text)
				}
			}
		}

		pptStr := strings.TrimSpace(pptText.String())
		pptxStr := strings.TrimSpace(pptxTexts[i+1])

		if pptStr == "" && pptxStr == "" {
			continue
		}

		// Compare by removing whitespace for fuzzy match
		pptClean := strings.ReplaceAll(strings.ReplaceAll(pptStr, " ", ""), "\n", "")
		pptxClean := strings.ReplaceAll(strings.ReplaceAll(pptxStr, " ", ""), "\n", "")

		if pptClean != pptxClean {
			mismatches++
			if mismatches <= 5 {
				fmt.Printf("Slide %d: TEXT MISMATCH\n", i+1)
				if len(pptClean) > 100 {
					fmt.Printf("  PPT:  %s...\n", pptClean[:100])
				} else {
					fmt.Printf("  PPT:  %s\n", pptClean)
				}
				if len(pptxClean) > 100 {
					fmt.Printf("  PPTX: %s...\n", pptxClean[:100])
				} else {
					fmt.Printf("  PPTX: %s\n", pptxClean)
				}
			}
		}
	}

	fmt.Printf("\n%d/%d slides have text mismatches\n", mismatches, len(slides))
}

func extractXMLText(xmlStr string) string {
	decoder := xml.NewDecoder(strings.NewReader(xmlStr))
	var text strings.Builder
	inT := false
	for {
		tok, err := decoder.Token()
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
				text.Write(t)
			}
		}
	}
	return text.String()
}
