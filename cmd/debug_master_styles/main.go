package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	presentation, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	slides := presentation.GetSlides()
	masters := presentation.GetMasters()

	// Check master default text styles
	fmt.Println("=== Master Default Text Styles ===")
	for ref, m := range masters {
		hasStyles := false
		for _, s := range m.DefaultTextStyles {
			if s.FontSize > 0 {
				hasStyles = true
				break
			}
		}
		if hasStyles {
			fmt.Printf("Master %d:\n", ref)
			for i, s := range m.DefaultTextStyles {
				if s.FontSize > 0 || s.FontName != "" {
					fmt.Printf("  Level %d: size=%d font=%q bold=%v color=%q\n",
						i, s.FontSize, s.FontName, s.Bold, s.Color)
				}
			}
		}
	}

	// Check how many runs still have fontSize=0 after master style application
	fmt.Println("\n=== Remaining fontSize=0 after master styles ===")
	totalRuns := 0
	zeroRuns := 0
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					totalRuns++
					if r.FontSize == 0 {
						zeroRuns++
					}
				}
			}
		}
	}
	fmt.Printf("Total runs: %d, fontSize=0: %d (%.1f%%)\n", totalRuns, zeroRuns, float64(zeroRuns)*100/float64(totalRuns))

	// Check PPTX for small font sizes
	fmt.Println("\n=== PPTX font size distribution ===")
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer zr.Close()

	szCounts := make(map[string]int)
	for _, f := range zr.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") || strings.Contains(f.Name, "_rels") {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		content := string(data)

		// Count sz values
		idx := 0
		for {
			pos := strings.Index(content[idx:], `sz="`)
			if pos < 0 {
				break
			}
			pos += idx + 4
			end := strings.Index(content[pos:], `"`)
			if end < 0 {
				break
			}
			sz := content[pos : pos+end]
			szCounts[sz]++
			idx = pos + end
		}
	}

	// Print sorted by count
	type szEntry struct {
		sz    string
		count int
	}
	var entries []szEntry
	for sz, count := range szCounts {
		entries = append(entries, szEntry{sz, count})
	}
	// Sort by count descending
	for i := 0; i < len(entries); i++ {
		for j := i + 1; j < len(entries); j++ {
			if entries[j].count > entries[i].count {
				entries[i], entries[j] = entries[j], entries[i]
			}
		}
	}
	for _, e := range entries {
		fmt.Printf("  sz=%s: %d\n", e.sz, e.count)
	}
}
