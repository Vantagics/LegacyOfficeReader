package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	f, _ := zip.OpenReader("testfie/test.pptx")
	defer f.Close()

	// Check slides 2, 3, 7, 35, 48 (TOC pages)
	targets := []string{"ppt/slides/slide2.xml", "ppt/slides/slide3.xml", "ppt/slides/slide7.xml", "ppt/slides/slide35.xml", "ppt/slides/slide48.xml"}
	for _, target := range targets {
		for _, zf := range f.File {
			if zf.Name == target {
				rc, _ := zf.Open()
				data, _ := io.ReadAll(rc)
				rc.Close()
				xml := string(data)
				
				fmt.Printf("=== %s ===\n", target)
				// Count custGeom vs prstGeom
				custCount := strings.Count(xml, "custGeom")
				prstCount := strings.Count(xml, "prstGeom")
				fmt.Printf("  custGeom: %d, prstGeom: %d\n", custCount, prstCount)
				
				// Check for line dash
				dashCount := strings.Count(xml, "prstDash")
				fmt.Printf("  prstDash: %d\n", dashCount)
				
				// Check for hexagon shapes (6 vertices = hexagon number shapes)
				// Look for shapes with text "1", "2", "3", "4"
				for _, num := range []string{"1", "2", "3", "4"} {
					idx := strings.Index(xml, ">"+num+"</a:t>")
					if idx >= 0 {
						// Find the enclosing sp element
						start := strings.LastIndex(xml[:idx], "<p:sp>")
						end := strings.Index(xml[idx:], "</p:sp>")
						if start >= 0 && end >= 0 {
							snippet := xml[start : idx+end+len("</p:sp>")]
							hasCustGeom := strings.Contains(snippet, "custGeom")
							hasSolidFill := strings.Contains(snippet, "solidFill")
							hasLine := strings.Contains(snippet, "<a:ln")
							fmt.Printf("  Number '%s': custGeom=%v solidFill=%v hasLine=%v\n", num, hasCustGeom, hasSolidFill, hasLine)
							
							// Extract line info
							lineIdx := strings.Index(snippet, "<a:ln")
							if lineIdx >= 0 {
								lineEnd := strings.Index(snippet[lineIdx:], "</a:ln>")
								if lineEnd >= 0 {
									fmt.Printf("    Line: %s\n", snippet[lineIdx:lineIdx+lineEnd+len("</a:ln>")])
								}
							}
						}
					}
				}
				
				// Check for underline shapes (text shapes with bottom-only line)
				fmt.Println()
			}
		}
	}
	
	// Also dump slide2 full XML for inspection
	for _, zf := range f.File {
		if zf.Name == "ppt/slides/slide2.xml" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			
			// Write to file for inspection
			os.WriteFile("slide2_output.xml", data, 0644)
			fmt.Println("Wrote slide2_output.xml for inspection")
		}
	}
}
