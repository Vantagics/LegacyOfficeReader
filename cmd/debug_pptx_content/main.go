package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to open PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// 1. Validate all XML files
	fmt.Println("=== XML Validation ===")
	xmlErrors := 0
	fileCount := 0
	for _, f := range zr.File {
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		fileCount++
		rc, err := f.Open()
		if err != nil {
			fmt.Printf("  ERROR opening %s: %v\n", f.Name, err)
			xmlErrors++
			continue
		}
		d := xml.NewDecoder(rc)
		for {
			_, err := d.Token()
			if err != nil {
				if err == io.EOF {
					break
				}
				fmt.Printf("  XML ERROR in %s: %v\n", f.Name, err)
				xmlErrors++
				break
			}
		}
		rc.Close()
	}
	fmt.Printf("  %d XML/rels files validated, %d errors\n", fileCount, xmlErrors)

	// 2. Check for sz="0"
	fmt.Println("\n=== Font Size Check ===")
	szZeroCount := 0
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		szZeroCount += strings.Count(string(data), `sz="0"`)
	}
	fmt.Printf("  sz=\"0\" occurrences: %d\n", szZeroCount)

	// 3. Check bullet usage
	fmt.Println("\n=== Bullet Check ===")
	buCharCount := 0
	buNoneCount := 0
	defaultBulletCount := 0
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)
		buCharCount += strings.Count(content, "buChar")
		buNoneCount += strings.Count(content, "buNone")
		defaultBulletCount += strings.Count(content, `buChar char="•"`)
	}
	fmt.Printf("  buChar: %d (default •: %d), buNone: %d\n", buCharCount, defaultBulletCount, buNoneCount)

	// 4. Check slide count matches
	fmt.Println("\n=== Slide Count ===")
	slideCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			slideCount++
		}
	}
	fmt.Printf("  Slides: %d (expected: 71)\n", slideCount)

	// 5. Check image count
	imageCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			imageCount++
		}
	}
	fmt.Printf("  Images: %d (expected: 166)\n", imageCount)

	// 6. Check layout count
	layoutCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			layoutCount++
		}
	}
	fmt.Printf("  Layouts: %d (expected: 7)\n", layoutCount)

	// 7. Font size distribution
	fmt.Println("\n=== Font Size Distribution ===")
	szRe := regexp.MustCompile(`sz="(\d+)"`)
	szDist := make(map[string]int)
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		for _, m := range szRe.FindAllStringSubmatch(string(data), -1) {
			szDist[m[1]]++
		}
	}
	for _, sz := range []string{"600", "700", "800", "900", "1000", "1100", "1200", "1400", "1600", "1800", "2000", "2200", "2400", "2800", "3200", "4000", "6000", "8300"} {
		if count, ok := szDist[sz]; ok {
			fmt.Printf("  sz=%s: %d\n", sz, count)
		}
	}

	// 8. Check normAutofit
	fmt.Println("\n=== Auto-fit ===")
	totalAutofit := 0
	totalBodyPr := 0
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)
		totalAutofit += strings.Count(content, "normAutofit")
		totalBodyPr += strings.Count(content, "<a:bodyPr")
	}
	fmt.Printf("  normAutofit: %d / %d bodyPr\n", totalAutofit, totalBodyPr)

	// 9. Check showMasterSp
	fmt.Println("\n=== showMasterSp ===")
	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if strings.Contains(content, `showMasterSp="1"`) {
				fmt.Println("  slide1: showMasterSp=1 (correct)")
			} else {
				fmt.Println("  slide1: MISSING showMasterSp=1")
			}
		}
	}
	for _, f := range zr.File {
		if f.Name == "ppt/slideLayouts/slideLayout1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if strings.Contains(content, `showMasterSp="0"`) {
				fmt.Println("  layout1: showMasterSp=0 (correct)")
			} else {
				fmt.Println("  layout1: MISSING showMasterSp=0")
			}
		}
	}

	// 10. Summary
	fmt.Println("\n=== Summary ===")
	if xmlErrors == 0 && szZeroCount == 0 && slideCount == 71 && imageCount == 166 {
		fmt.Println("  ALL CHECKS PASSED")
	} else {
		fmt.Println("  SOME CHECKS FAILED")
	}
}
