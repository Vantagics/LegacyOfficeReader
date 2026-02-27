package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	issues := 0

	// Check each slide
	for slideNum := 1; slideNum <= 71; slideNum++ {
		fname := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		for _, f := range zr.File {
			if f.Name == fname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Check 1: showMasterSp="1" present
				if !strings.Contains(content, `showMasterSp="1"`) {
					fmt.Printf("ISSUE: Slide %d missing showMasterSp=\"1\"\n", slideNum)
					issues++
				}

				// Check 2: No empty text runs (sz="0")
				if strings.Contains(content, `sz="0"`) {
					fmt.Printf("ISSUE: Slide %d has sz=\"0\" (zero font size)\n", slideNum)
					issues++
				}

				// Check 3: Check for negative positions or sizes
				d := xml.NewDecoder(strings.NewReader(content))
				for {
					tok, err := d.Token()
					if err != nil {
						break
					}
					if se, ok := tok.(xml.StartElement); ok {
						if se.Name.Local == "ext" {
							for _, attr := range se.Attr {
								if attr.Name.Local == "cx" || attr.Name.Local == "cy" {
									val, _ := strconv.Atoi(attr.Value)
									if val < 0 {
										fmt.Printf("ISSUE: Slide %d has negative %s=%d\n", slideNum, attr.Name.Local, val)
										issues++
									}
								}
							}
						}
					}
				}

				// Check 4: Verify text content is not empty for text shapes
				textCount := strings.Count(content, "<a:t>") + strings.Count(content, `<a:t `)
				if textCount == 0 {
					// Some slides may legitimately have no text (image-only slides)
					spCount := strings.Count(content, `txBox="1"`)
					if spCount > 0 {
						fmt.Printf("WARNING: Slide %d has %d textboxes but no text content\n", slideNum, spCount)
					}
				}

				break
			}
		}

		// Check slide rels
		relsFname := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)
		for _, f := range zr.File {
			if f.Name == relsFname {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Check: has slideLayout reference
				if !strings.Contains(content, "slideLayout") {
					fmt.Printf("ISSUE: Slide %d rels missing slideLayout reference\n", slideNum)
					issues++
				}
				break
			}
		}
	}

	// Check all layouts have proper rels
	for i := 1; i <= 7; i++ {
		relsFname := fmt.Sprintf("ppt/slideLayouts/_rels/slideLayout%d.xml.rels", i)
		found := false
		for _, f := range zr.File {
			if f.Name == relsFname {
				found = true
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)
				if !strings.Contains(content, "slideMaster") {
					fmt.Printf("ISSUE: Layout %d rels missing slideMaster reference\n", i)
					issues++
				}
				break
			}
		}
		if !found {
			fmt.Printf("ISSUE: Layout %d rels file missing\n", i)
			issues++
		}
	}

	// Check presentation.xml
	for _, f := range zr.File {
		if f.Name == "ppt/presentation.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			slideRefs := strings.Count(content, "<p:sldId")
			fmt.Printf("Presentation: %d slide references\n", slideRefs)
			if slideRefs != 71 {
				fmt.Printf("ISSUE: Expected 71 slide references, got %d\n", slideRefs)
				issues++
			}

			// Check slide size
			if !strings.Contains(content, `cx="12192000"`) || !strings.Contains(content, `cy="6858000"`) {
				fmt.Printf("ISSUE: Slide size mismatch\n")
				issues++
			}
			break
		}
	}

	// Check theme
	for _, f := range zr.File {
		if f.Name == "ppt/theme/theme1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if !strings.Contains(content, "clrScheme") {
				fmt.Printf("ISSUE: Theme missing color scheme\n")
				issues++
			}
			break
		}
	}

	fmt.Printf("\nTotal issues: %d\n", issues)
}
