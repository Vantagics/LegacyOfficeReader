package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// List all files
	fmt.Println("=== PPTX Files ===")
	slideCount := 0
	layoutCount := 0
	imageCount := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideCount++
		}
		if strings.HasPrefix(f.Name, "ppt/slideLayouts/") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			layoutCount++
		}
		if strings.HasPrefix(f.Name, "ppt/media/") {
			imageCount++
		}
	}
	fmt.Printf("Slides: %d, Layouts: %d, Images: %d\n", slideCount, layoutCount, imageCount)

	// Check slide 1 XML for watermark/layout issues
	for si := 1; si <= 5; si++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
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
				showMaster := "?"
				if strings.Contains(content, `showMasterSp="1"`) {
					showMaster = "1"
				} else if strings.Contains(content, `showMasterSp="0"`) {
					showMaster = "0"
				}

				fmt.Printf("\nSlide %d: shapes=%d pics=%d connectors=%d hasBg=%v showMaster=%s\n",
					si, spCount, picCount, cxnCount, hasBg, showMaster)

				// Check for color issues - look for srgbClr values
				colorIdx := 0
				pos := 0
				for colorIdx < 10 {
					idx := strings.Index(content[pos:], `srgbClr val="`)
					if idx < 0 {
						break
					}
					pos += idx + 13
					end := strings.Index(content[pos:], `"`)
					if end < 0 {
						break
					}
					color := content[pos : pos+end]
					// Find surrounding context
					ctxStart := pos - 100
					if ctxStart < 0 {
						ctxStart = 0
					}
					ctxEnd := pos + end + 50
					if ctxEnd > len(content) {
						ctxEnd = len(content)
					}
					ctx := content[ctxStart:ctxEnd]
					ctx = strings.ReplaceAll(ctx, "\n", " ")
					if len(ctx) > 200 {
						ctx = ctx[:200]
					}
					fmt.Printf("  Color[%d]: %s ctx=...%s...\n", colorIdx, color, ctx)
					pos += end
					colorIdx++
				}
			}
		}
	}

	// Check layout 1 XML
	for li := 1; li <= 3; li++ {
		name := fmt.Sprintf("ppt/slideLayouts/slideLayout%d.xml", li)
		for _, f := range zr.File {
			if f.Name == name {
				rc, _ := f.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				content := string(data)

				spCount := strings.Count(content, "<p:sp>") + strings.Count(content, "<p:sp ")
				picCount := strings.Count(content, "<p:pic>") + strings.Count(content, "<p:pic ")
				hasBg := strings.Contains(content, "<p:bg>")
				hasGradFill := strings.Contains(content, "gradFill")
				hasTitleBg := strings.Contains(content, "TitleBg") || strings.Contains(content, "titleBg")

				fmt.Printf("\nLayout %d: shapes=%d pics=%d hasBg=%v hasGradFill=%v hasTitleBg=%v\n",
					li, spCount, picCount, hasBg, hasGradFill, hasTitleBg)

				if len(content) > 3000 {
					content = content[:3000] + "..."
				}
				fmt.Printf("  Content: %s\n", content)
			}
		}
	}
}
