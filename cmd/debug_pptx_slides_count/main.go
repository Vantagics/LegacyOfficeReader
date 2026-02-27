package main

import (
	"archive/zip"
	"fmt"
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

	slideFiles := 0
	slideRels := 0
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") && !strings.Contains(f.Name, "_rels") {
			slideFiles++
		}
		if strings.HasPrefix(f.Name, "ppt/slides/_rels/slide") && strings.HasSuffix(f.Name, ".xml.rels") {
			slideRels++
		}
	}
	fmt.Printf("Slide XML files: %d\n", slideFiles)
	fmt.Printf("Slide rels files: %d\n", slideRels)

	// Check content types
	for _, f := range zr.File {
		if f.Name == "[Content_Types].xml" {
			rc, _ := f.Open()
			data := make([]byte, f.UncompressedSize64)
			rc.Read(data)
			rc.Close()
			content := string(data)
			slideOverrides := strings.Count(content, "presentationml.slide+xml")
			fmt.Printf("Content types slide overrides: %d\n", slideOverrides)
			break
		}
	}

	// Check presentation.xml slide refs
	for _, f := range zr.File {
		if f.Name == "ppt/presentation.xml" {
			rc, _ := f.Open()
			data := make([]byte, f.UncompressedSize64)
			rc.Read(data)
			rc.Close()
			content := string(data)
			slideRefs := strings.Count(content, "<p:sldId")
			fmt.Printf("Presentation slide refs: %d\n", slideRefs)
			break
		}
	}

	// Check presentation rels
	for _, f := range zr.File {
		if f.Name == "ppt/_rels/presentation.xml.rels" {
			rc, _ := f.Open()
			data := make([]byte, f.UncompressedSize64)
			rc.Read(data)
			rc.Close()
			content := string(data)
			slideRels2 := strings.Count(content, "relationships/slide\"")
			fmt.Printf("Presentation rels slide refs: %d\n", slideRels2)
			break
		}
	}
}
