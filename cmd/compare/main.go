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
		fmt.Fprintf(os.Stderr, "Error opening PPT: %v\n", err)
		os.Exit(1)
	}

	pptSlides := p.GetSlides()
	fmt.Printf("PPT: %d slides\n", len(pptSlides))

	// Open PPTX
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	pptxSlideCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			pptxSlideCount++
		}
	}
	fmt.Printf("PPTX: %d slides\n", pptxSlideCount)

	if len(pptSlides) != pptxSlideCount {
		fmt.Printf("MISMATCH: PPT has %d slides, PPTX has %d slides\n", len(pptSlides), pptxSlideCount)
	}

	// Compare text content per slide
	mismatches := 0
	for i := 0; i < len(pptSlides) && i < pptxSlideCount; i++ {
		pptTexts := pptSlides[i].GetTexts()
		pptxTexts := getPptxSlideTexts(&zr.Reader, i+1)

		// Combine all PPT texts
		pptAll := strings.Join(pptTexts, " ")
		pptxAll := strings.Join(pptxTexts, " ")

		// Normalize whitespace
		pptAll = normalizeWS(pptAll)
		pptxAll = normalizeWS(pptxAll)

		if pptAll != pptxAll {
			// Check if PPTX contains all PPT text
			missing := false
			for _, t := range pptTexts {
				t = normalizeWS(t)
				if t == "" {
					continue
				}
				if !strings.Contains(pptxAll, t) {
					// Try shorter substring
					if len(t) > 20 {
						t = t[:20]
					}
					if !strings.Contains(pptxAll, t) {
						fmt.Printf("Slide %d: PPT text not found in PPTX: %q\n", i+1, truncate(t, 60))
						missing = true
					}
				}
			}
			if missing {
				mismatches++
			}
		}
	}

	// Check shapes per slide
	fmt.Printf("\n=== Shape comparison ===\n")
	for i := 0; i < len(pptSlides) && i < pptxSlideCount; i++ {
		shapes := pptSlides[i].GetShapes()
		pptxShapes := countPptxShapes(&zr.Reader, i+1)
		if len(shapes) != pptxShapes {
			fmt.Printf("Slide %d: PPT shapes=%d, PPTX shapes=%d\n", i+1, len(shapes), pptxShapes)
		}
	}

	// Check images
	pptImages := p.GetImages()
	pptxImages := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			pptxImages++
		}
	}
	fmt.Printf("\nPPT images: %d, PPTX images: %d\n", len(pptImages), pptxImages)

	if mismatches > 0 {
		fmt.Printf("\n%d slides have text mismatches\n", mismatches)
	} else {
		fmt.Printf("\nAll slide texts match!\n")
	}
}

func getPptxSlideTexts(zr *zip.Reader, slideNum int) []string {
	name := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
	for _, f := range zr.File {
		if f.Name == name {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			return extractTexts(data)
		}
	}
	return nil
}

func countPptxShapes(zr *zip.Reader, slideNum int) int {
	name := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
	for _, f := range zr.File {
		if f.Name == name {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Count <p:sp>, <p:pic>, <p:cxnSp>
			count := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ") +
				strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ") +
				strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
			return count
		}
	}
	return 0
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

func normalizeWS(s string) string {
	s = strings.ReplaceAll(s, "\r\n", " ")
	s = strings.ReplaceAll(s, "\r", " ")
	s = strings.ReplaceAll(s, "\n", " ")
	s = strings.ReplaceAll(s, "\v", " ")
	s = strings.ReplaceAll(s, "\t", " ")
	for strings.Contains(s, "  ") {
		s = strings.ReplaceAll(s, "  ", " ")
	}
	return strings.TrimSpace(s)
}

func truncate(s string, n int) string {
	if len(s) > n {
		return s[:n] + "..."
	}
	return s
}
