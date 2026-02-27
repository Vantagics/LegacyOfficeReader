// +build ignore

package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io/ioutil"
	"os"
	"regexp"
	"strconv"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

var issues []string

func addIssue(format string, args ...interface{}) {
	issues = append(issues, fmt.Sprintf(format, args...))
}

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open ppt: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	images := p.GetImages()
	fonts := p.GetFonts()
	slideW, slideH := p.GetSlideSize()

	fmt.Println("Reading PPTX...")
	pptxData, err := ioutil.ReadFile("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "read pptx: %v\n", err)
		os.Exit(1)
	}

	zr, err := zip.NewReader(strings.NewReader(string(pptxData)), int64(len(pptxData)))
	if err != nil {
		fmt.Fprintf(os.Stderr, "open zip: %v\n", err)
		os.Exit(1)
	}

	fileSet := make(map[string]bool)
	for _, f := range zr.File {
		fileSet[f.Name] = true
	}

	// Count PPTX slides
	pptxSlides := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			pptxSlides++
		}
	}

	// === ROUND 1: Structure ===
	fmt.Println("\n=== ROUND 1: Structure ===")
	fmt.Printf("PPT slides: %d, PPTX slides: %d\n", len(slides), pptxSlides)
	if len(slides) != pptxSlides {
		addIssue("SLIDE COUNT MISMATCH: PPT=%d PPTX=%d", len(slides), pptxSlides)
	}

	// Check required files
	required := []string{
		"[Content_Types].xml", "_rels/.rels",
		"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels",
		"ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout1.xml",
		"ppt/theme/theme1.xml", "ppt/presProps.xml", "ppt/viewProps.xml", "ppt/tableStyles.xml",
	}
	for _, r := range required {
		if !fileSet[r] {
			addIssue("MISSING FILE: %s", r)
		}
	}

	// Check slide size
	presXML := readZipFile(zr, "ppt/presentation.xml")
	if presXML != nil {
		sldSzRe := regexp.MustCompile(`<p:sldSz cx="(\d+)" cy="(\d+)"`)
		m := sldSzRe.FindSubmatch(presXML)
		if m != nil {
			cx, _ := strconv.Atoi(string(m[1]))
			cy, _ := strconv.Atoi(string(m[2]))
			if int32(cx) != slideW || int32(cy) != slideH {
				addIssue("SLIDE SIZE: PPT=%dx%d PPTX=%dx%d", slideW, slideH, cx, cy)
			}
			fmt.Printf("Slide size: PPT=%dx%d PPTX=%dx%d\n", slideW, slideH, cx, cy)
		}
	}

	// === ROUND 2: XML validity ===
	fmt.Println("\n=== ROUND 2: XML validity ===")
	xmlErrors := 0
	for _, f := range zr.File {
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		data := readZipFileEntry(f)
		if data == nil {
			continue
		}
		d := xml.NewDecoder(strings.NewReader(string(data)))
		for {
			_, err := d.Token()
			if err != nil {
				if err.Error() != "EOF" {
					xmlErrors++
					if xmlErrors <= 3 {
						addIssue("XML ERROR in %s: %v", f.Name, err)
					}
				}
				break
			}
		}
	}
	fmt.Printf("XML parse errors: %d\n", xmlErrors)

	// === ROUND 3: Relationships ===
	fmt.Println("\n=== ROUND 3: Relationships ===")
	dupRels := 0
	missingTargets := 0
	for i := 1; i <= pptxSlides; i++ {
		relsPath := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", i)
		relsData := readZipFile(zr, relsPath)
		if relsData == nil {
			addIssue("MISSING RELS: %s", relsPath)
			continue
		}
		idRe := regexp.MustCompile(`Id="([^"]*)"`)
		matches := idRe.FindAllSubmatch(relsData, -1)
		seen := make(map[string]bool)
		for _, m := range matches {
			id := string(m[1])
			if seen[id] {
				dupRels++
			}
			seen[id] = true
		}
		targetRe := regexp.MustCompile(`Target="([^"]*)"`)
		tMatches := targetRe.FindAllSubmatch(relsData, -1)
		for _, m := range tMatches {
			target := string(m[1])
			if strings.Contains(target, "media/") {
				fullPath := "ppt/" + strings.TrimPrefix(target, "../")
				if !fileSet[fullPath] {
					missingTargets++
				}
			}
		}
	}
	fmt.Printf("Duplicate rel IDs: %d, Missing targets: %d\n", dupRels, missingTargets)
	if dupRels > 0 {
		addIssue("DUPLICATE REL IDS: %d", dupRels)
	}
	if missingTargets > 0 {
		addIssue("MISSING TARGETS: %d", missingTargets)
	}

	// === ROUND 4: Shape properties ===
	fmt.Println("\n=== ROUND 4: Shape properties ===")
	negDims := 0
	emptyParas := 0
	emptyLn := 0
	pptxTotalSp, pptxTotalPic, pptxTotalCxn := 0, 0, 0

	for i := 1; i <= pptxSlides; i++ {
		slideXML := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i))
		if slideXML == nil {
			continue
		}
		content := string(slideXML)
		pptxTotalSp += strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
		pptxTotalPic += strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
		pptxTotalCxn += strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")

		extRe := regexp.MustCompile(`<a:ext cx="(-?\d+)" cy="(-?\d+)"`)
		for _, m := range extRe.FindAllStringSubmatch(content, -1) {
			cx, _ := strconv.Atoi(m[1])
			cy, _ := strconv.Atoi(m[2])
			if cx < 0 || cy < 0 {
				negDims++
			}
		}
		emptyPRe := regexp.MustCompile(`<a:p>((?:<a:pPr[^/]*/a:pPr>)?)</a:p>`)
		emptyParas += len(emptyPRe.FindAllString(content, -1))
		lnRe := regexp.MustCompile(`<a:ln w="[^"]*"></a:ln>`)
		emptyLn += len(lnRe.FindAllString(content, -1))
	}
	totalPPTXShapes := pptxTotalSp + pptxTotalPic + pptxTotalCxn
	fmt.Printf("PPTX shapes: %d (sp=%d pic=%d cxn=%d)\n", totalPPTXShapes, pptxTotalSp, pptxTotalPic, pptxTotalCxn)
	fmt.Printf("Negative dims: %d, Empty paragraphs: %d, Empty <a:ln>: %d\n", negDims, emptyParas, emptyLn)
	if negDims > 0 {
		addIssue("NEGATIVE DIMENSIONS: %d", negDims)
	}

	// === ROUND 5: Per-slide 1:1 comparison ===
	fmt.Println("\n=== ROUND 5: Per-slide comparison ===")
	shapeMismatch := 0
	textMismatch := 0
	maxSlides := len(slides)
	if pptxSlides < maxSlides {
		maxSlides = pptxSlides
	}

	pptFonts := make(map[string]int)
	pptBold, pptItalic, pptColor, pptSize := 0, 0, 0, 0
	pptAlign, pptBullet, pptSpacing, pptRot := 0, 0, 0, 0
	pptShapes, pptImgShapes, pptTextShapes := 0, 0, 0

	for si := 0; si < maxSlides; si++ {
		pptSlide := slides[si]
		pptShapeList := pptSlide.GetShapes()
		pptShapeCount := len(pptShapeList)

		// PPT stats for this slide
		for _, sh := range pptShapeList {
			pptShapes++
			if sh.IsImage && sh.ImageIdx >= 0 && sh.ImageIdx < len(images) {
				pptImgShapes++
			}
			if sh.IsText || len(sh.Paragraphs) > 0 {
				pptTextShapes++
			}
			if sh.Rotation != 0 {
				pptRot++
			}
			for _, para := range sh.Paragraphs {
				if para.Alignment > 0 {
					pptAlign++
				}
				if para.HasBullet {
					pptBullet++
				}
				if para.SpaceBefore != 0 || para.SpaceAfter != 0 || para.LineSpacing != 0 {
					pptSpacing++
				}
				for _, run := range para.Runs {
					if run.FontName != "" {
						pptFonts[run.FontName]++
					}
					if run.Bold {
						pptBold++
					}
					if run.Italic {
						pptItalic++
					}
					if run.Color != "" {
						pptColor++
					}
					if run.FontSize > 0 {
						pptSize++
					}
				}
			}
		}

		// PPTX stats for this slide
		slideXML := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", si+1))
		if slideXML == nil {
			continue
		}
		content := string(slideXML)
		xSp := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
		xPic := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
		xCxn := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
		xTotal := xSp + xPic + xCxn

		if pptShapeCount != xTotal {
			shapeMismatch++
			if shapeMismatch <= 10 {
				fmt.Printf("  Slide %d: PPT shapes=%d PPTX shapes=%d (sp=%d pic=%d cxn=%d)\n",
					si+1, pptShapeCount, xTotal, xSp, xPic, xCxn)
			}
		}

		// Compare text content
		var pptTexts []string
		for _, sh := range pptShapeList {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t != "" {
						pptTexts = append(pptTexts, t)
					}
				}
			}
		}

		var pptxTexts []string
		d := xml.NewDecoder(strings.NewReader(content))
		inT := false
		for {
			tok, err := d.Token()
			if err != nil {
				break
			}
			switch t := tok.(type) {
			case xml.StartElement:
				if t.Name.Local == "t" && t.Name.Space == "http://schemas.openxmlformats.org/drawingml/2006/main" {
					inT = true
				}
			case xml.CharData:
				if inT {
					txt := strings.TrimSpace(string(t))
					if txt != "" {
						pptxTexts = append(pptxTexts, txt)
					}
				}
			case xml.EndElement:
				if t.Name.Local == "t" {
					inT = false
				}
			}
		}

		pptTextJoined := strings.Join(pptTexts, "|")
		pptxTextJoined := strings.Join(pptxTexts, "|")
		if pptTextJoined != pptxTextJoined {
			textMismatch++
			if textMismatch <= 5 {
				pptSnip := pptTextJoined
				pptxSnip := pptxTextJoined
				if len(pptSnip) > 80 {
					pptSnip = pptSnip[:80] + "..."
				}
				if len(pptxSnip) > 80 {
					pptxSnip = pptxSnip[:80] + "..."
				}
				fmt.Printf("  Slide %d text diff:\n    PPT:  %s\n    PPTX: %s\n", si+1, pptSnip, pptxSnip)
			}
		}
	}
	fmt.Printf("Shape count mismatches: %d/%d slides\n", shapeMismatch, maxSlides)
	fmt.Printf("Text content mismatches: %d/%d slides\n", textMismatch, maxSlides)

	// === ROUND 6: Formatting stats ===
	fmt.Println("\n=== ROUND 6: Formatting comparison ===")
	fmt.Printf("PPT: %d shapes (%d img, %d text)\n", pptShapes, pptImgShapes, pptTextShapes)
	fmt.Printf("PPT: bold=%d italic=%d color=%d size=%d\n", pptBold, pptItalic, pptColor, pptSize)
	fmt.Printf("PPT: align=%d bullet=%d spacing=%d rot=%d\n", pptAlign, pptBullet, pptSpacing, pptRot)
	fmt.Printf("PPT fonts: %v\n", pptFonts)
	fmt.Printf("PPT font collection: %v\n", fonts)
	fmt.Printf("PPT images: %d\n", len(images))

	// PPTX formatting stats
	pptxBold, pptxItalic, pptxColor, pptxSize := 0, 0, 0, 0
	pptxAlign, pptxBullet, pptxRot := 0, 0, 0
	pptxFonts := make(map[string]int)

	for i := 1; i <= pptxSlides; i++ {
		slideXML := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i))
		if slideXML == nil {
			continue
		}
		d := xml.NewDecoder(strings.NewReader(string(slideXML)))
		inRPr := false
		inRun := false
		for {
			tok, err := d.Token()
			if err != nil {
				break
			}
			switch t := tok.(type) {
			case xml.StartElement:
				if t.Name.Local == "r" && t.Name.Space == "http://schemas.openxmlformats.org/drawingml/2006/main" {
					inRun = true
				}
				if t.Name.Local == "rPr" && t.Name.Space == "http://schemas.openxmlformats.org/drawingml/2006/main" {
					inRPr = true
					for _, attr := range t.Attr {
						switch attr.Name.Local {
						case "b":
							if attr.Value == "1" {
								pptxBold++
							}
						case "i":
							if attr.Value == "1" {
								pptxItalic++
							}
						case "sz":
							if attr.Value != "" && attr.Value != "0" {
								pptxSize++
							}
						}
					}
				}
				if inRPr && t.Name.Local == "srgbClr" {
					pptxColor++
				}
				if inRun && t.Name.Local == "latin" {
					for _, attr := range t.Attr {
						if attr.Name.Local == "typeface" && attr.Value != "" {
							pptxFonts[attr.Value]++
						}
					}
				}
				if t.Name.Local == "xfrm" {
					for _, attr := range t.Attr {
						if attr.Name.Local == "rot" && attr.Value != "" && attr.Value != "0" {
							pptxRot++
						}
					}
				}
				if t.Name.Local == "pPr" && t.Name.Space == "http://schemas.openxmlformats.org/drawingml/2006/main" {
					for _, attr := range t.Attr {
						if attr.Name.Local == "algn" && attr.Value != "" && attr.Value != "l" {
							pptxAlign++
						}
					}
				}
				if t.Name.Local == "buChar" {
					pptxBullet++
				}
			case xml.EndElement:
				if t.Name.Local == "rPr" {
					inRPr = false
				}
				if t.Name.Local == "r" && t.Name.Space == "http://schemas.openxmlformats.org/drawingml/2006/main" {
					inRun = false
				}
			}
		}
	}
	fmt.Printf("PPTX: bold=%d italic=%d color=%d size=%d\n", pptxBold, pptxItalic, pptxColor, pptxSize)
	fmt.Printf("PPTX: align=%d bullet=%d rot=%d\n", pptxAlign, pptxBullet, pptxRot)
	fmt.Printf("PPTX fonts: %v\n", pptxFonts)

	// === ROUND 7: Delta analysis ===
	fmt.Println("\n=== ROUND 7: Delta analysis ===")
	fmt.Printf("bold:   PPT=%d PPTX=%d delta=%d\n", pptBold, pptxBold, pptxBold-pptBold)
	fmt.Printf("italic: PPT=%d PPTX=%d delta=%d\n", pptItalic, pptxItalic, pptxItalic-pptItalic)
	fmt.Printf("color:  PPT=%d PPTX=%d delta=%d\n", pptColor, pptxColor, pptxColor-pptColor)
	fmt.Printf("size:   PPT=%d PPTX=%d delta=%d\n", pptSize, pptxSize, pptxSize-pptSize)
	fmt.Printf("align:  PPT=%d PPTX=%d delta=%d\n", pptAlign, pptxAlign, pptxAlign-pptAlign)
	fmt.Printf("bullet: PPT=%d PPTX=%d delta=%d\n", pptBullet, pptxBullet, pptxBullet-pptBullet)
	fmt.Printf("rot:    PPT=%d PPTX=%d delta=%d\n", pptRot, pptxRot, pptxRot-pptRot)
	fmt.Printf("shapes: PPT=%d PPTX=%d delta=%d\n", pptShapes, totalPPTXShapes, totalPPTXShapes-pptShapes)

	// Font deltas
	allFontNames := make(map[string]bool)
	for k := range pptFonts {
		allFontNames[k] = true
	}
	for k := range pptxFonts {
		allFontNames[k] = true
	}
	for fn := range allFontNames {
		pc := pptFonts[fn]
		xc := pptxFonts[fn]
		if pc != xc {
			fmt.Printf("  font %s: PPT=%d PPTX=%d delta=%d\n", fn, pc, xc, xc-pc)
		}
	}

	// === SUMMARY ===
	fmt.Printf("\n========== SUMMARY ==========\n")
	fmt.Printf("PPT: %d slides, %d shapes, %d images\n", len(slides), pptShapes, len(images))
	fmt.Printf("PPTX: %d slides, %d shapes, %d images\n", pptxSlides, totalPPTXShapes, pptxTotalPic)
	if len(issues) == 0 {
		fmt.Println("\n*** ALL CHECKS PASSED ***")
	} else {
		fmt.Printf("\n*** %d ISSUES FOUND ***\n", len(issues))
		for i, issue := range issues {
			if i >= 30 {
				fmt.Printf("  ... and %d more\n", len(issues)-30)
				break
			}
			fmt.Printf("  %s\n", issue)
		}
	}
}

func readZipFile(zr *zip.Reader, name string) []byte {
	for _, f := range zr.File {
		if f.Name == name {
			return readZipFileEntry(f)
		}
	}
	return nil
}

func readZipFileEntry(f *zip.File) []byte {
	rc, err := f.Open()
	if err != nil {
		return nil
	}
	defer rc.Close()
	data, err := ioutil.ReadAll(rc)
	if err != nil {
		return nil
	}
	return data
}

// suppress unused imports
var _ = ppt.Presentation{}
