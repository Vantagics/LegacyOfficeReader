package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
)

type slideInfo struct {
	shapes     int
	pics       int
	connectors int
	texts      []string
	hasBg      bool
	bgType     string // "solid", "blip", "none"
}

func analyzeSlide(data []byte) slideInfo {
	content := string(data)
	si := slideInfo{}
	si.shapes = strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
	si.pics = strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
	si.connectors = strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
	si.hasBg = strings.Contains(content, "<p:bg>")
	if strings.Contains(content, "blipFill") && si.hasBg {
		si.bgType = "blip"
	} else if strings.Contains(content, "solidFill") && si.hasBg {
		si.bgType = "solid"
	} else {
		si.bgType = "none"
	}

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
					si.texts = append(si.texts, s)
				}
			}
		}
	}
	return si
}

func main() {
	// Compare reference.pptx with test.pptx
	refFile := "testfie/reference.pptx"
	testFile := "testfie/test.pptx"

	ref, err := zip.OpenReader(refFile)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Cannot open reference: %v\n", err)
		os.Exit(1)
	}
	defer ref.Close()

	test, err := zip.OpenReader(testFile)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Cannot open test: %v\n", err)
		os.Exit(1)
	}
	defer test.Close()

	// Count slides in each
	refSlides := 0
	testSlides := 0
	for _, f := range ref.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			refSlides++
		}
	}
	for _, f := range test.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			testSlides++
		}
	}

	fmt.Printf("Reference PPTX: %d slides\n", refSlides)
	fmt.Printf("Test PPTX: %d slides\n", testSlides)

	// Compare each slide
	maxSlides := refSlides
	if testSlides > maxSlides {
		maxSlides = testSlides
	}

	for i := 1; i <= maxSlides; i++ {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", i)

		var refInfo, testInfo slideInfo
		var refFound, testFound bool

		for _, f := range ref.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				refInfo = analyzeSlide(data)
				refFound = true
			}
		}
		for _, f := range test.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				testInfo = analyzeSlide(data)
				testFound = true
			}
		}

		if !refFound && !testFound {
			continue
		}

		// Check for differences
		hasDiff := false
		if refInfo.hasBg != testInfo.hasBg || refInfo.bgType != testInfo.bgType {
			hasDiff = true
		}
		if len(refInfo.texts) != len(testInfo.texts) {
			hasDiff = true
		}
		if refInfo.shapes != testInfo.shapes || refInfo.pics != testInfo.pics {
			hasDiff = true
		}

		if hasDiff || !refFound || !testFound {
			fmt.Printf("\nSlide %d DIFF:\n", i)
			if refFound {
				firstText := ""
				if len(refInfo.texts) > 0 {
					firstText = refInfo.texts[0]
					if len(firstText) > 40 {
						firstText = firstText[:40] + "..."
					}
				}
				fmt.Printf("  REF:  shapes=%d pics=%d conn=%d texts=%d bg=%v(%s) first=%q\n",
					refInfo.shapes, refInfo.pics, refInfo.connectors, len(refInfo.texts), refInfo.hasBg, refInfo.bgType, firstText)
			} else {
				fmt.Printf("  REF:  [missing]\n")
			}
			if testFound {
				firstText := ""
				if len(testInfo.texts) > 0 {
					firstText = testInfo.texts[0]
					if len(firstText) > 40 {
						firstText = firstText[:40] + "..."
					}
				}
				fmt.Printf("  TEST: shapes=%d pics=%d conn=%d texts=%d bg=%v(%s) first=%q\n",
					testInfo.shapes, testInfo.pics, testInfo.connectors, len(testInfo.texts), testInfo.hasBg, testInfo.bgType, firstText)
			} else {
				fmt.Printf("  TEST: [missing]\n")
			}
		}
	}
}
