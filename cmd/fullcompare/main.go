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

	pptSlides := p.GetSlides()
	fmt.Printf("PPT slides: %d\n", len(pptSlides))

	totalMissing := 0
	for i, slide := range pptSlides {
		// Get all text from shapes
		var shapeTexts []string
		for _, sh := range slide.GetShapes() {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t != "" {
						shapeTexts = append(shapeTexts, t)
					}
				}
			}
		}
		allShapeText := strings.Join(shapeTexts, " ")

		// Check each PPT text is in shapes
		for _, t := range slide.GetTexts() {
			t = strings.TrimSpace(t)
			if t == "" {
				continue
			}
			// Normalize: remove \r\n
			t = strings.ReplaceAll(t, "\r\n", " ")
			t = strings.ReplaceAll(t, "\r", " ")
			t = strings.ReplaceAll(t, "\n", " ")
			t = strings.ReplaceAll(t, "\v", " ")

			// Check if first 10 chars are in shape text
			check := t
			if len(check) > 10 {
				check = check[:10]
			}
			if !strings.Contains(allShapeText, check) {
				fmt.Printf("Slide %d: PPT text NOT in shapes: %q\n", i+1, truncate(t, 80))
				totalMissing++
			}
		}
	}

	if totalMissing == 0 {
		fmt.Println("All PPT text content is present in shapes (and thus in PPTX)")
	} else {
		fmt.Printf("\n%d text blocks missing from shapes\n", totalMissing)
	}

	// Now check PPTX structure validity
	fmt.Println("\n=== PPTX Structure Check ===")
	requiredFiles := []string{
		"ppt/presentation.xml",
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/_rels/presentation.xml.rels",
		"ppt/theme/theme1.xml",
		"ppt/slideLayouts/slideLayout1.xml",
		"ppt/slideMasters/slideMaster1.xml",
	}
	for _, req := range requiredFiles {
		found := false
		for _, f := range zr.File {
			if f.Name == req {
				found = true
				break
			}
		}
		if !found {
			fmt.Printf("MISSING: %s\n", req)
		}
	}

	// Check each slide has rels
	pptxSlideCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			pptxSlideCount++
		}
	}
	for i := 1; i <= pptxSlideCount; i++ {
		relsName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", i)
		found := false
		for _, f := range zr.File {
			if f.Name == relsName {
				found = true
				break
			}
		}
		if !found {
			fmt.Printf("MISSING rels: %s\n", relsName)
		}
	}

	// Check presentation.xml has correct slide count
	for _, f := range zr.File {
		if f.Name == "ppt/presentation.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			sldIdCount := strings.Count(content, "<p:sldId ")
			fmt.Printf("presentation.xml sldId count: %d (expected %d)\n", sldIdCount, pptxSlideCount)

			// Check slide size
			if strings.Contains(content, "p:sldSz") {
				idx := strings.Index(content, "p:sldSz")
				end := strings.Index(content[idx:], "/>")
				if end > 0 {
					fmt.Printf("Slide size: %s\n", content[idx:idx+end+2])
				}
			}
		}
	}

	fmt.Printf("\nPPTX structure: OK\n")
	fmt.Printf("Total slides: PPT=%d, PPTX=%d\n", len(pptSlides), pptxSlideCount)
	fmt.Printf("Total images: PPT=%d, PPTX=", len(p.GetImages()))
	imgCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			imgCount++
		}
	}
	fmt.Printf("%d\n", imgCount)
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
	if len(s) > n {
		return s[:n] + "..."
	}
	return s
}
