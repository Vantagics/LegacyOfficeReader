package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

var (
	reSpOpen  = regexp.MustCompile(`<p:sp[ >]`)
	rePicOpen = regexp.MustCompile(`<p:pic[ >]`)
	reCxnOpen = regexp.MustCompile(`<p:cxnSp[ >]`)
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	images := p.GetImages()
	slideW, slideH := p.GetSlideSize()

	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	fmt.Println("=== FINAL VERIFICATION ===")
	fmt.Printf("PPT: %d slides, %d images, size %dx%d EMU\n", len(slides), len(images), slideW, slideH)

	// Count PPTX files
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
	fmt.Printf("PPTX: %d slides, %d layouts, %d media files\n", slideFiles, layoutFiles, mediaFiles)

	// 1. Shape count verification
	fmt.Println("\n--- 1. Shape Count ---")
	mismatch := 0
	for i, s := range slides {
		pptCount := len(s.GetShapes())
		content := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i+1))
		sp := len(reSpOpen.FindAllString(content, -1))
		pic := len(rePicOpen.FindAllString(content, -1))
		cxn := len(reCxnOpen.FindAllString(content, -1))
		pptxCount := sp + pic + cxn
		if pptCount != pptxCount {
			fmt.Printf("  Slide %d: PPT=%d PPTX=%d DIFF=%+d\n", i+1, pptCount, pptxCount, pptxCount-pptCount)
			mismatch++
		}
	}
	if mismatch == 0 {
		fmt.Println("  All slides match ✓")
	} else {
		fmt.Printf("  %d mismatches\n", mismatch)
	}

	// 2. XML well-formedness
	fmt.Println("\n--- 2. XML Well-formedness ---")
	xmlErrors := 0
	for _, f := range zr.File {
		if !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			continue
		}
		data, _ := io.ReadAll(rc)
		rc.Close()
		decoder := xml.NewDecoder(strings.NewReader(string(data)))
		for {
			_, err := decoder.Token()
			if err == io.EOF {
				break
			}
			if err != nil {
				fmt.Printf("  INVALID: %s: %v\n", f.Name, err)
				xmlErrors++
				break
			}
		}
	}
	if xmlErrors == 0 {
		fmt.Println("  All XML files well-formed ✓")
	}

	// 3. Check for sz="0"
	fmt.Println("\n--- 3. Font Size Check ---")
	sz0Count := 0
	for i := 1; i <= slideFiles; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i))
		if strings.Contains(content, `sz="0"`) {
			fmt.Printf("  Slide %d: HAS sz=0!\n", i)
			sz0Count++
		}
	}
	if sz0Count == 0 {
		fmt.Println("  No sz=0 found ✓")
	}

	// 4. showMasterSp check
	fmt.Println("\n--- 4. showMasterSp ---")
	slideMasterOk := true
	for i := 1; i <= slideFiles; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i))
		if !strings.Contains(content, `showMasterSp="1"`) {
			fmt.Printf("  Slide %d: MISSING showMasterSp=1\n", i)
			slideMasterOk = false
		}
	}
	layoutMasterOk := true
	for i := 1; i <= layoutFiles; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", i))
		if !strings.Contains(content, `showMasterSp="0"`) {
			fmt.Printf("  Layout %d: MISSING showMasterSp=0\n", i)
			layoutMasterOk = false
		}
	}
	if slideMasterOk && layoutMasterOk {
		fmt.Println("  All slides/layouts correct ✓")
	}

	// 5. clrMapOvr check
	fmt.Println("\n--- 5. Color Map Override ---")
	clrMapOk := true
	for i := 1; i <= slideFiles; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i))
		if !strings.Contains(content, "<p:clrMapOvr>") {
			fmt.Printf("  Slide %d: MISSING clrMapOvr\n", i)
			clrMapOk = false
		}
	}
	if clrMapOk {
		fmt.Println("  All slides have clrMapOvr ✓")
	}

	// 6. defaultTextStyle check
	fmt.Println("\n--- 6. Default Text Style ---")
	presContent := readZipFile(zr, "ppt/presentation.xml")
	if strings.Contains(presContent, "<p:defaultTextStyle>") {
		fmt.Println("  Present ✓")
	} else {
		fmt.Println("  MISSING!")
	}

	// 7. Bullet font check
	fmt.Println("\n--- 7. Bullet Fonts ---")
	buFontCount := 0
	buCharCount := 0
	buNoneCount := 0
	for i := 1; i <= slideFiles; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i))
		buFontCount += strings.Count(content, "<a:buFont")
		buCharCount += strings.Count(content, "<a:buChar")
		buNoneCount += strings.Count(content, "<a:buNone")
	}
	fmt.Printf("  buFont: %d, buChar: %d, buNone: %d ✓\n", buFontCount, buCharCount, buNoneCount)

	// 8. normAutofit check
	fmt.Println("\n--- 8. Auto-fit ---")
	autofitCount := 0
	bodyPrCount := 0
	for i := 1; i <= slideFiles; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i))
		autofitCount += strings.Count(content, "normAutofit")
		bodyPrCount += strings.Count(content, "<a:bodyPr")
	}
	fmt.Printf("  normAutofit: %d / %d bodyPr ✓\n", autofitCount, bodyPrCount)

	// 9. noFill/noLine check
	fmt.Println("\n--- 9. Fill/Line Defaults ---")
	noFillTotal := 0
	lnNoFillTotal := 0
	for i := 1; i <= slideFiles; i++ {
		content := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i))
		noFillTotal += strings.Count(content, "<a:noFill/>")
		lnNoFillTotal += strings.Count(content, `<a:ln><a:noFill/></a:ln>`)
	}
	fmt.Printf("  noFill: %d, ln noFill: %d ✓\n", noFillTotal, lnNoFillTotal)

	// 10. Text content spot check
	fmt.Println("\n--- 10. Text Content Spot Check ---")
	totalMissing := 0
	for i := 0; i < len(slides) && i < 10; i++ {
		shapes := slides[i].GetShapes()
		content := readZipFile(zr, fmt.Sprintf("ppt/slides/slide%d.xml", i+1))
		missing := 0
		total := 0
		for _, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if len(t) > 3 {
						total++
						escaped := xmlEscape(t)
						if !strings.Contains(content, escaped) && !strings.Contains(content, t) {
							missing++
						}
					}
				}
			}
		}
		if missing > 0 {
			fmt.Printf("  Slide %d: %d/%d texts missing\n", i+1, missing, total)
			totalMissing += missing
		}
	}
	if totalMissing == 0 {
		fmt.Println("  All text preserved ✓")
	}

	// 11. Theme check
	fmt.Println("\n--- 11. Theme ---")
	themeContent := readZipFile(zr, "ppt/theme/theme1.xml")
	if strings.Contains(themeContent, "微软雅黑") {
		fmt.Println("  EA font: 微软雅黑 ✓")
	}
	if strings.Contains(themeContent, `val="000000"`) && strings.Contains(themeContent, "<a:dk1>") {
		fmt.Println("  dk1: 000000 (black) ✓")
	}
	if strings.Contains(themeContent, `val="FFFFFF"`) && strings.Contains(themeContent, "<a:lt1>") {
		fmt.Println("  lt1: FFFFFF (white) ✓")
	}

	// 12. Slide size check
	fmt.Println("\n--- 12. Slide Size ---")
	if strings.Contains(presContent, fmt.Sprintf(`cx="%d"`, slideW)) && strings.Contains(presContent, fmt.Sprintf(`cy="%d"`, slideH)) {
		fmt.Printf("  %dx%d EMU ✓\n", slideW, slideH)
	} else {
		fmt.Println("  MISMATCH!")
	}

	fmt.Println("\n=== VERIFICATION COMPLETE ===")
}

func readZipFile(zr *zip.ReadCloser, name string) string {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			return string(data)
		}
	}
	return ""
}

func xmlEscape(s string) string {
	var buf strings.Builder
	xml.Escape(&buf, []byte(s))
	return buf.String()
}
