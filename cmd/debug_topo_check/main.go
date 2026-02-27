package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Find paragraphs around topology diagrams
			// Search for "部署拓扑图" and "典型部署"
			keywords := []string{"部署拓扑图", "典型部署", "高级威胁检测", "本地威胁发现", "文件威胁检测"}
			for _, kw := range keywords {
				idx := strings.Index(content, kw)
				if idx >= 0 {
					start := idx - 200
					if start < 0 { start = 0 }
					end := idx + 200
					if end > len(content) { end = len(content) }
					fmt.Printf("\n=== Found '%s' at offset %d ===\n", kw, idx)
					fmt.Println(content[start:end])
				}
			}
			
			// Count drawings and check their types
			fmt.Println("\n=== Drawing analysis ===")
			parts := strings.Split(content, "<w:drawing>")
			for i := 1; i < len(parts); i++ {
				end := strings.Index(parts[i], "</w:drawing>")
				if end < 0 { continue }
				drawing := parts[i][:end]
				
				isAnchor := strings.Contains(drawing, "<wp:anchor")
				isInline := strings.Contains(drawing, "<wp:inline")
				
				// Get extent
				extIdx := strings.Index(drawing, "<wp:extent")
				extent := ""
				if extIdx >= 0 {
					extEnd := strings.Index(drawing[extIdx:], "/>")
					if extEnd >= 0 {
						extent = drawing[extIdx:extIdx+extEnd+2]
					}
				}
				
				// Get image ref
				blipIdx := strings.Index(drawing, "r:embed=")
				blipRef := ""
				if blipIdx >= 0 {
					blipRef = drawing[blipIdx:blipIdx+20]
				}
				
				wrapType := "none"
				if strings.Contains(drawing, "wrapTopAndBottom") {
					wrapType = "topAndBottom"
				} else if strings.Contains(drawing, "wrapNone") {
					wrapType = "none"
				} else if strings.Contains(drawing, "wrapSquare") {
					wrapType = "square"
				}
				
				fmt.Printf("Drawing[%d]: anchor=%v inline=%v wrap=%s %s %s\n", i, isAnchor, isInline, wrapType, extent, blipRef)
			}
		}
	}
}
