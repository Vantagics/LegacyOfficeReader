package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Parse original PPT
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPT: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	images := p.GetImages()
	masters := p.GetMasters()
	slideW, slideH := p.GetSlideSize()

	fmt.Printf("=== PPT Summary ===\n")
	fmt.Printf("Slides: %d, Images: %d, Masters: %d\n", len(slides), len(images), len(masters))
	fmt.Printf("Slide size: %d x %d EMU\n", slideW, slideH)

	// Analyze each slide
	fmt.Printf("\n=== Per-Slide Analysis ===\n")
	for i, s := range slides {
		shapes := s.GetShapes()
		bg := s.GetBackground()
		texts := s.GetTexts()

		imgCount := 0
		textCount := 0
		connCount := 0
		for _, sh := range shapes {
			if sh.IsImage {
				imgCount++
			} else if sh.ShapeType == 20 || (sh.ShapeType >= 32 && sh.ShapeType <= 40) {
				connCount++
			} else if sh.IsText || len(sh.Paragraphs) > 0 {
				textCount++
			}
		}

		// Get first text snippet
		snippet := ""
		for _, t := range texts {
			t = strings.TrimSpace(t)
			if len(t) > 0 {
				if len(t) > 40 {
					snippet = t[:40] + "..."
				} else {
					snippet = t
				}
				break
			}
		}

		bgInfo := "none"
		if bg.HasBackground {
			if bg.ImageIdx >= 0 {
				bgInfo = fmt.Sprintf("image(%d)", bg.ImageIdx)
			} else if bg.FillColor != "" {
				bgInfo = fmt.Sprintf("solid(%s)", bg.FillColor)
			}
		}

		fmt.Printf("Slide %2d: shapes=%2d (img=%d txt=%d conn=%d) bg=%s masterRef=%d  %s\n",
			i+1, len(shapes), imgCount, textCount, connCount, bgInfo, s.GetMasterRef(), snippet)
	}

	// Analyze PPTX output
	fmt.Printf("\n=== PPTX Output Analysis ===\n")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	slideFiles := 0
	layoutFiles := 0
	mediaFiles := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideFiles++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			layoutFiles++
		}
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaFiles++
		}
	}
	fmt.Printf("Slides: %d, Layouts: %d, Media: %d\n", slideFiles, layoutFiles, mediaFiles)

	// Check each slide XML for common issues
	fmt.Printf("\n=== PPTX Slide Content Check ===\n")
	for i := 1; i <= slideFiles; i++ {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		content := readZipFile(zr, fname)
		if content == "" {
			fmt.Printf("Slide %d: MISSING\n", i)
			continue
		}

		// Count shapes
		spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
		picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
		cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
		hasBg := strings.Contains(content, "<p:bg>")
		hasShowMaster := strings.Contains(content, `showMasterSp="1"`)

		// Check for sz="0"
		hasSz0 := strings.Contains(content, `sz="0"`)

		// Check for empty text bodies
		emptyTxBody := strings.Count(content, "<p:txBody><a:bodyPr")

		issues := []string{}
		if hasSz0 {
			issues = append(issues, "HAS sz=0!")
		}
		if !hasShowMaster {
			issues = append(issues, "MISSING showMasterSp")
		}

		issueStr := ""
		if len(issues) > 0 {
			issueStr = " ISSUES: " + strings.Join(issues, ", ")
		}

		fmt.Printf("Slide %2d: sp=%2d pic=%2d cxn=%2d bg=%v txBody=%d%s\n",
			i, spCount, picCount, cxnCount, hasBg, emptyTxBody, issueStr)
	}

	// Check layouts
	fmt.Printf("\n=== Layout Analysis ===\n")
	for i := 1; i <= layoutFiles; i++ {
		fname := fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", i)
		content := readZipFile(zr, fname)
		if content == "" {
			continue
		}
		spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
		picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
		cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
		hasBg := strings.Contains(content, "<p:bg>")
		hasShowMaster0 := strings.Contains(content, `showMasterSp="0"`)

		issues := []string{}
		if !hasShowMaster0 {
			issues = append(issues, "MISSING showMasterSp=0")
		}
		issueStr := ""
		if len(issues) > 0 {
			issueStr = " ISSUES: " + strings.Join(issues, ", ")
		}

		fmt.Printf("Layout %d: sp=%d pic=%d cxn=%d bg=%v%s\n",
			i, spCount, picCount, cxnCount, hasBg, issueStr)
	}

	// Detailed check: compare PPT shapes vs PPTX shapes per slide
	fmt.Printf("\n=== Shape Count Comparison (PPT vs PPTX) ===\n")
	mismatchCount := 0
	for i, s := range slides {
		pptShapes := len(s.GetShapes())
		slideNum := i + 1
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		content := readZipFile(zr, fname)
		pptxShapes := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ") +
			strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ") +
			strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")

		if pptShapes != pptxShapes {
			fmt.Printf("Slide %2d: PPT=%2d PPTX=%2d  MISMATCH\n", slideNum, pptShapes, pptxShapes)
			mismatchCount++
		}
	}
	if mismatchCount == 0 {
		fmt.Printf("All %d slides have matching shape counts ✓\n", len(slides))
	} else {
		fmt.Printf("%d slides have mismatched shape counts\n", mismatchCount)
	}

	// Check for text content preservation
	fmt.Printf("\n=== Text Content Spot Check (first 5 slides) ===\n")
	for i := 0; i < 5 && i < len(slides); i++ {
		shapes := slides[i].GetShapes()
		slideNum := i + 1
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		content := readZipFile(zr, fname)

		// Extract text from PPT
		var pptTexts []string
		for _, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t != "" {
						pptTexts = append(pptTexts, t)
					}
				}
			}
		}

		// Check if PPT texts appear in PPTX
		missing := 0
		for _, t := range pptTexts {
			escaped := xmlEscape(t)
			if !strings.Contains(content, escaped) && !strings.Contains(content, t) {
				if len(t) > 3 { // skip very short texts
					missing++
				}
			}
		}

		fmt.Printf("Slide %d: %d text runs, %d missing in PPTX\n", slideNum, len(pptTexts), missing)
	}

	// Check for font size distribution in PPTX
	fmt.Printf("\n=== Font Size Distribution in PPTX ===\n")
	szDist := make(map[string]int)
	for i := 1; i <= slideFiles; i++ {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		content := readZipFile(zr, fname)
		// Find all sz="..." occurrences
		for {
			idx := strings.Index(content, `sz="`)
			if idx < 0 {
				break
			}
			content = content[idx+4:]
			end := strings.Index(content, `"`)
			if end < 0 {
				break
			}
			sz := content[:end]
			szDist[sz]++
			content = content[end:]
		}
	}
	for sz, count := range szDist {
		fmt.Printf("  sz=%s: %d\n", sz, count)
	}

	// Check for paragraph spacing issues
	fmt.Printf("\n=== Paragraph Spacing Check ===\n")
	lnSpcCount := 0
	spcBefCount := 0
	spcAftCount := 0
	for i := 1; i <= slideFiles; i++ {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", i)
		content := readZipFile(zr, fname)
		lnSpcCount += strings.Count(content, "<a:lnSpc>")
		spcBefCount += strings.Count(content, "<a:spcBef>")
		spcAftCount += strings.Count(content, "<a:spcAft>")
	}
	fmt.Printf("Line spacing elements: %d\n", lnSpcCount)
	fmt.Printf("Space before elements: %d\n", spcBefCount)
	fmt.Printf("Space after elements: %d\n", spcAftCount)

	// Check PPT paragraph spacing data
	fmt.Printf("\n=== PPT Paragraph Spacing Data ===\n")
	pptLnSpc := 0
	pptSpcBef := 0
	pptSpcAft := 0
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				if para.LineSpacing != 0 {
					pptLnSpc++
				}
				if para.SpaceBefore != 0 {
					pptSpcBef++
				}
				if para.SpaceAfter != 0 {
					pptSpcAft++
				}
			}
		}
	}
	fmt.Printf("PPT paragraphs with line spacing: %d\n", pptLnSpc)
	fmt.Printf("PPT paragraphs with space before: %d\n", pptSpcBef)
	fmt.Printf("PPT paragraphs with space after: %d\n", pptSpcAft)
}

func readZipFile(zr *zip.ReadCloser, name string) string {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			buf := make([]byte, f.UncompressedSize64)
			n, _ := rc.Read(buf)
			return string(buf[:n])
		}
	}
	return ""
}

func xmlEscape(s string) string {
	var buf strings.Builder
	xml.Escape(&buf, []byte(s))
	return buf.String()
}
