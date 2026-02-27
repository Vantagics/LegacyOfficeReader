package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer zr.Close()

	// Check a few slides for shape count and content
	for slideNum := 1; slideNum <= 5; slideNum++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				// Count shapes
				spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
				picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
				cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
				hasBg := strings.Contains(content, "<p:bg>")
				hasClrMap := strings.Contains(content, "clrMapOvr")
				hasMasterSp := strings.Contains(content, "showMasterSp")

				fmt.Printf("Slide %d: sp=%d pic=%d cxn=%d bg=%v clrMap=%v masterSp=%v size=%d\n",
					slideNum, spCount, picCount, cxnCount, hasBg, hasClrMap, hasMasterSp, len(data))

				// Check for sz="0"
				if strings.Contains(content, `sz="0"`) {
					fmt.Printf("  WARNING: contains sz=\"0\"\n")
				}

				// Check for normAutofit
				autofitCount := strings.Count(content, "normAutofit")
				bodyPrCount := strings.Count(content, "bodyPr")
				fmt.Printf("  bodyPr=%d normAutofit=%d\n", bodyPrCount, autofitCount)
			}
		}
	}

	// Check layout files
	fmt.Println("\n=== Layouts ===")
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
			picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
			cxnCount := strings.Count(content, "<p:cxnSp>") + strings.Count(content, "<p:cxnSp ")
			hasBg := strings.Contains(content, "<p:bg>")
			hasShowMaster := strings.Contains(content, `showMasterSp="0"`)

			fmt.Printf("%s: sp=%d pic=%d cxn=%d bg=%v showMasterSp0=%v size=%d\n",
				f.Name, spCount, picCount, cxnCount, hasBg, hasShowMaster, len(data))
		}
	}

	// Check theme
	fmt.Println("\n=== Theme ===")
	for _, f := range zr.File {
		if f.Name == "ppt/theme/theme1.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			// Extract dk1 and lt1
			if idx := strings.Index(content, "dk1"); idx >= 0 {
				end := idx + 80
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("dk1: %s\n", content[idx:end])
			}
			if idx := strings.Index(content, "lt1"); idx >= 0 {
				end := idx + 80
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("lt1: %s\n", content[idx:end])
			}
			// Check font
			if idx := strings.Index(content, "微软雅黑"); idx >= 0 {
				fmt.Println("Theme has 微软雅黑 font")
			}
		}
	}

	// Check presentation.xml
	fmt.Println("\n=== Presentation ===")
	for _, f := range zr.File {
		if f.Name == "ppt/presentation.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			hasDefaultTextStyle := strings.Contains(content, "defaultTextStyle")
			hasSaveSubset := strings.Contains(content, "saveSubsetFonts")
			fmt.Printf("defaultTextStyle=%v saveSubsetFonts=%v size=%d\n", hasDefaultTextStyle, hasSaveSubset, len(data))
			// Extract slide size
			if idx := strings.Index(content, "sldSz"); idx >= 0 {
				end := idx + 60
				if end > len(content) {
					end = len(content)
				}
				fmt.Printf("sldSz: %s\n", content[idx:end])
			}
		}
	}
}
