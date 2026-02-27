package main

import (
	"archive/zip"
	"fmt"
	"os"
	"regexp"
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

	colorRe := regexp.MustCompile(`<a:srgbClr val="([A-F0-9]{6})"`)
	colorMap := make(map[string]int)

	for _, file := range f.File {
		if !strings.HasPrefix(file.Name, "ppt/") || !strings.HasSuffix(file.Name, ".xml") {
			continue
		}
		rc, _ := file.Open()
		buf := make([]byte, file.UncompressedSize64)
		n, _ := rc.Read(buf)
		rc.Close()
		content := string(buf[:n])

		matches := colorRe.FindAllStringSubmatch(content, -1)
		for _, m := range matches {
			colorMap[m[1]]++
		}
	}

	// Sort by frequency
	type colorFreq struct {
		color string
		count int
	}
	var sorted []colorFreq
	for c, n := range colorMap {
		sorted = append(sorted, colorFreq{c, n})
	}
	sort.Slice(sorted, func(i, j int) bool { return sorted[i].count > sorted[j].count })

	fmt.Printf("=== Color Distribution (all slides + layouts) ===\n")
	fmt.Printf("Total unique colors: %d\n\n", len(sorted))
	for _, cf := range sorted {
		label := ""
		switch cf.color {
		case "FFFFFF":
			label = " (white)"
		case "000000":
			label = " (black)"
		case "0C0D0E":
			label = " (dk1 - dark text)"
		case "FF0000":
			label = " (red)"
		case "4472C4":
			label = " (blue accent)"
		case "0C80DE":
			label = " (bright blue)"
		case "E9EBF5":
			label = " (light purple bg)"
		case "CFD5EA":
			label = " (medium purple bg)"
		case "003296":
			label = " (dark blue line)"
		case "06213C":
			label = " (very dark blue)"
		case "FFD966":
			label = " (gold/yellow)"
		case "42A5F5":
			label = " (light blue accent)"
		case "2196F3":
			label = " (material blue)"
		}
		fmt.Printf("  #%s: %4d%s\n", cf.color, cf.count, label)
	}
}
