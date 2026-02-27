package main

import (
	"archive/zip"
	"fmt"
	"os"
	"sort"
	"strings"
)

func main() {
	f, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	// Categorize files
	categories := make(map[string]int)
	totalSize := int64(0)
	var xmlFiles []string

	for _, file := range f.File {
		totalSize += int64(file.UncompressedSize64)
		parts := strings.Split(file.Name, "/")
		if len(parts) > 1 {
			categories[parts[0]+"/"+parts[1]]++
		} else {
			categories[parts[0]]++
		}
		if strings.HasSuffix(file.Name, ".xml") || strings.HasSuffix(file.Name, ".rels") {
			xmlFiles = append(xmlFiles, file.Name)
		}
	}

	fmt.Printf("=== PPTX Structure ===\n")
	fmt.Printf("Total files: %d\n", len(f.File))
	fmt.Printf("Total uncompressed size: %.1f MB\n\n", float64(totalSize)/1024/1024)

	// Sort categories
	var cats []string
	for c := range categories {
		cats = append(cats, c)
	}
	sort.Strings(cats)
	for _, c := range cats {
		fmt.Printf("  %-40s %d files\n", c, categories[c])
	}

	// Check required PPTX files
	fmt.Printf("\n=== Required Files Check ===\n")
	required := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/_rels/presentation.xml.rels",
		"ppt/theme/theme1.xml",
		"ppt/slideMasters/slideMaster1.xml",
		"ppt/slideMasters/_rels/slideMaster1.xml.rels",
		"ppt/presProps.xml",
		"ppt/viewProps.xml",
		"ppt/tableStyles.xml",
	}

	fileSet := make(map[string]bool)
	for _, file := range f.File {
		fileSet[file.Name] = true
	}

	allPresent := true
	for _, req := range required {
		if fileSet[req] {
			fmt.Printf("  ✓ %s\n", req)
		} else {
			fmt.Printf("  ✗ MISSING: %s\n", req)
			allPresent = false
		}
	}

	// Check slide/layout/rels consistency
	fmt.Printf("\n=== Consistency Check ===\n")
	slideCount := 0
	layoutCount := 0
	slideRelsCount := 0
	layoutRelsCount := 0
	for _, file := range f.File {
		if strings.HasPrefix(file.Name, "ppt/slides/slide") && strings.HasSuffix(file.Name, ".xml") {
			slideCount++
		}
		if strings.HasPrefix(file.Name, "ppt/slideLayouts/slideLayout") && strings.HasSuffix(file.Name, ".xml") {
			layoutCount++
		}
		if strings.HasPrefix(file.Name, "ppt/slides/_rels/slide") && strings.HasSuffix(file.Name, ".xml.rels") {
			slideRelsCount++
		}
		if strings.HasPrefix(file.Name, "ppt/slideLayouts/_rels/slideLayout") && strings.HasSuffix(file.Name, ".xml.rels") {
			layoutRelsCount++
		}
	}

	fmt.Printf("  Slides: %d (rels: %d) %s\n", slideCount, slideRelsCount,
		map[bool]string{true: "✓", false: "✗ MISMATCH"}[slideCount == slideRelsCount])
	fmt.Printf("  Layouts: %d (rels: %d) %s\n", layoutCount, layoutRelsCount,
		map[bool]string{true: "✓", false: "✗ MISMATCH"}[layoutCount == layoutRelsCount])

	if allPresent && slideCount == slideRelsCount && layoutCount == layoutRelsCount {
		fmt.Printf("\n✓ PPTX structure is valid\n")
	}
}
