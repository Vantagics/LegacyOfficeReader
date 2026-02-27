package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	var docXML string
	for _, f := range zr.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			docXML = string(data)
		}
	}

	// Find all image references and their context
	fmt.Printf("=== Image Placements ===\n")
	for i := 1; i <= 8; i++ {
		ref := fmt.Sprintf(`r:embed="rImg%d"`, i)
		idx := 0
		for {
			pos := strings.Index(docXML[idx:], ref)
			if pos < 0 {
				break
			}
			absPos := idx + pos
			// Show context: find the enclosing <w:p> 
			pStart := strings.LastIndex(docXML[:absPos], "<w:p>")
			if pStart < 0 {
				pStart = absPos - 200
			}
			// Find extent
			extStart := strings.LastIndex(docXML[pStart:absPos], `<wp:extent`)
			if extStart >= 0 {
				extEnd := strings.Index(docXML[pStart+extStart:], "/>")
				if extEnd >= 0 {
					fmt.Printf("Image %d: %s\n", i, docXML[pStart+extStart:pStart+extStart+extEnd+2])
				}
			}
			idx = absPos + len(ref)
		}
	}

	// Check for the "引言" heading and surrounding content
	fmt.Printf("\n=== Content around '引言' ===\n")
	idx := strings.Index(docXML, "引言")
	if idx >= 0 {
		start := idx - 200
		if start < 0 {
			start = 0
		}
		end := idx + 500
		if end > len(docXML) {
			end = len(docXML)
		}
		fmt.Println(docXML[start:end])
	}

	// Check table structure
	fmt.Printf("\n=== Table Structure ===\n")
	tblIdx := strings.Index(docXML, "<w:tbl>")
	if tblIdx >= 0 {
		tblEnd := strings.Index(docXML[tblIdx:], "</w:tbl>")
		if tblEnd >= 0 {
			tblContent := docXML[tblIdx : tblIdx+tblEnd+8]
			fmt.Printf("Table length: %d chars\n", len(tblContent))
			// Count rows and cells
			fmt.Printf("Rows: %d\n", strings.Count(tblContent, "<w:tr>"))
			fmt.Printf("Cells: %d\n", strings.Count(tblContent, "<w:tc>"))
			// Show first 1000 chars
			if len(tblContent) > 1000 {
				fmt.Println(tblContent[:1000])
			} else {
				fmt.Println(tblContent)
			}
		}
	}
}
