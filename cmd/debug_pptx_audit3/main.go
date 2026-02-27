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

// Audit the generated PPTX against the source PPT for consistency issues
func main() {
	// Parse source PPT
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		panic(err)
	}
	defer f.Close()
	p, err := ppt.OpenReader(f)
	if err != nil {
		panic(err)
	}

	slides := p.GetSlides()
	fmt.Printf("=== PPT Source: %d slides ===\n\n", len(slides))

	// Check a few representative slides for content
	checkSlides := []int{0, 1, 2, 3, 4, 12, 20, 40, 50, 60, 70}
	for _, idx := range checkSlides {
		if idx >= len(slides) {
			continue
		}
		s := slides[idx]
		shapes := s.GetShapes()
		fmt.Printf("--- Slide %d: %d shapes ---\n", idx+1, len(shapes))
		textShapes := 0
		imageShapes := 0
		for si, sh := range shapes {
			if sh.IsImage {
				imageShapes++
			}
			if sh.IsText && len(sh.Paragraphs) > 0 {
				textShapes++
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						if strings.TrimSpace(run.Text) != "" {
							fmt.Printf("  Shape[%d] type=%d pos=(%d,%d) sz=(%d,%d) fontSize=%d font=%q color=%q text=%q\n",
								si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height,
								run.FontSize, run.FontName, run.Color, truncate(run.Text, 40))
						}
					}
				}
			}
		}
		fmt.Printf("  Summary: %d text shapes, %d image shapes\n\n", textShapes, imageShapes)
	}

	// Now check the generated PPTX
	fmt.Println("\n=== PPTX Output Analysis ===")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		panic(err)
	}
	defer zr.Close()

	// Check layout files
	layoutCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") {
			layoutCount++
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Count shapes in layout
			picCount := strings.Count(content, "<p:pic>")
			spCount := strings.Count(content, "<p:sp>")
			cxnCount := strings.Count(content, "<p:cxnSp>")
			hasBg := strings.Contains(content, "<p:bg>")
			showMaster := ""
			if strings.Contains(content, `showMasterSp="0"`) {
				showMaster = "showMasterSp=0"
			} else if strings.Contains(content, `showMasterSp="1"`) {
				showMaster = "showMasterSp=1"
			}

			fmt.Printf("Layout %s: %d pics, %d shapes, %d connectors, bg=%v %s\n",
				f.Name, picCount, spCount, cxnCount, hasBg, showMaster)
		}
	}

	// Check a few slide files for content
	for _, idx := range checkSlides {
		slideNum := idx + 1
		slideName := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		for _, f := range zr.File {
			if f.Name == slideName {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				picCount := strings.Count(content, "<p:pic>")
				spCount := strings.Count(content, "<p:sp>")
				cxnCount := strings.Count(content, "<p:cxnSp>")
				hasBg := strings.Contains(content, "<p:bg>")
				showMaster := ""
				if strings.Contains(content, `showMasterSp="1"`) {
					showMaster = "showMasterSp=1"
				} else if strings.Contains(content, `showMasterSp="0"`) {
					showMaster = "showMasterSp=0"
				}

				// Extract font sizes used
				fontSizes := extractAttrValues(content, "sz")
				// Extract layout reference
				layoutRef := ""
				relsName := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)
				for _, rf := range zr.File {
					if rf.Name == relsName {
						rrc, _ := rf.Open()
						rdata, _ := io.ReadAll(rrc)
						rrc.Close()
						if strings.Contains(string(rdata), "slideLayout") {
							start := strings.Index(string(rdata), "slideLayout")
							end := strings.Index(string(rdata)[start:], `"`)
							if end > 0 {
								layoutRef = string(rdata)[start : start+end]
							}
						}
					}
				}

				fmt.Printf("Slide %d: %d pics, %d shapes, %d connectors, bg=%v %s layout=%s fontSizes=%v\n",
					slideNum, picCount, spCount, cxnCount, hasBg, showMaster, layoutRef, fontSizes)
			}
		}
	}

	// Check for common issues
	fmt.Println("\n=== Issue Detection ===")

	// 1. Check if all slides have showMasterSp="1"
	missingShowMaster := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			if !strings.Contains(string(data), `showMasterSp="1"`) {
				missingShowMaster++
			}
		}
	}
	fmt.Printf("Slides missing showMasterSp=1: %d\n", missingShowMaster)

	// 2. Check layout rels for image references
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/_rels/") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			imgCount := strings.Count(content, "image")
			masterRef := strings.Contains(content, "slideMaster")
			fmt.Printf("Layout rels %s: %d image refs, masterRef=%v\n", f.Name, imgCount, masterRef)
		}
	}

	// 3. Check slideMaster for layout references
	for _, f := range zr.File {
		if f.Name == "ppt/slideMasters/slideMaster1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			layoutRefs := strings.Count(content, "slideLayout")
			fmt.Printf("SlideMaster: %d layout references\n", layoutRefs)
		}
	}

	// 4. Check for font size issues (sz=0 or very small)
	fmt.Println("\n=== Font Size Distribution ===")
	szDist := make(map[string]int)
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			sizes := extractAttrValues(string(data), "sz")
			for _, sz := range sizes {
				szDist[sz]++
			}
		}
	}
	for sz, count := range szDist {
		fmt.Printf("  sz=%s: %d occurrences\n", sz, count)
	}

	// 5. Check image count in media folder
	mediaCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/media/") {
			mediaCount++
		}
	}
	fmt.Printf("\nMedia files: %d\n", mediaCount)
}

func truncate(s string, n int) string {
	r := []rune(s)
	if len(r) > n {
		return string(r[:n]) + "..."
	}
	return s
}

func extractAttrValues(content, attr string) []string {
	var result []string
	seen := make(map[string]bool)
	search := attr + `="`
	idx := 0
	for {
		pos := strings.Index(content[idx:], search)
		if pos < 0 {
			break
		}
		start := idx + pos + len(search)
		end := strings.Index(content[start:], `"`)
		if end < 0 {
			break
		}
		val := content[start : start+end]
		if !seen[val] {
			seen[val] = true
			result = append(result, val)
		}
		idx = start + end + 1
	}
	// Deduplicate - just return unique values
	_ = xml.Name{} // suppress unused import
	return result
}
