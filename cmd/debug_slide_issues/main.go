package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	masters := p.GetMasters()

	// Check each slide for potential issues
	for i, slide := range slides {
		shapes := slide.GetShapes()
		bg := slide.GetBackground()
		masterRef := slide.GetMasterRef()
		master, hasMaster := masters[masterRef]

		// Determine if the master has a dark background
		masterHasDarkBg := false
		masterHasBgImage := false
		if hasMaster {
			if master.Background.ImageIdx >= 0 {
				masterHasBgImage = true
			}
			if master.Background.FillColor != "" {
				masterHasDarkBg = isDark(master.Background.FillColor)
			}
		}

		// Check for text color issues
		for si, sh := range shapes {
			if !sh.IsText || len(sh.Paragraphs) == 0 {
				continue
			}
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					if run.Color == "" && !sh.NoFill && sh.FillColor == "" {
						// Text with no explicit color on a shape with no fill
						// If master has dark bg image, text should be white
						if masterHasBgImage || masterHasDarkBg {
							fmt.Printf("ISSUE Slide %d Shape[%d] Para[%d] Run[%d]: no color, master has dark bg (img=%v), text=%q\n",
								i+1, si, pi, ri, masterHasBgImage, truncate(run.Text, 40))
						}
					}
					if run.FontSize == 0 {
						fmt.Printf("ISSUE Slide %d Shape[%d] Para[%d] Run[%d]: fontSize=0, text=%q\n",
							i+1, si, pi, ri, truncate(run.Text, 40))
					}
				}
			}
		}

		// Check for background issues
		if !bg.HasBackground && hasMaster && !master.Background.HasBackground {
			fmt.Printf("INFO Slide %d: no background, master also has no background\n", i+1)
		}
	}

	// Check generated PPTX
	fmt.Println("\n=== Checking generated PPTX ===")
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPTX: %v\n", err)
		return
	}
	defer r.Close()

	// Check slide1.xml for issues
	for _, f := range r.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Check for flowChartMultidocument (should be rare)
			if strings.Contains(content, "flowChartMultidocument") {
				count := strings.Count(content, "flowChartMultidocument")
				fmt.Printf("SHAPE %s: has %d flowChartMultidocument shapes\n", f.Name, count)
			}

			// Check for missing font sizes
			if strings.Contains(content, `sz="0"`) {
				count := strings.Count(content, `sz="0"`)
				fmt.Printf("FONT %s: has %d sz=\"0\" occurrences\n", f.Name, count)
			}
		}
	}
}

func isDark(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	r := hexVal(hex[0])*16 + hexVal(hex[1])
	g := hexVal(hex[2])*16 + hexVal(hex[3])
	b := hexVal(hex[4])*16 + hexVal(hex[5])
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	return lum < 128
}

func hexVal(c byte) int {
	if c >= '0' && c <= '9' {
		return int(c - '0')
	}
	if c >= 'a' && c <= 'f' {
		return int(c-'a') + 10
	}
	if c >= 'A' && c <= 'F' {
		return int(c-'A') + 10
	}
	return 0
}

func truncate(s string, n int) string {
	r := []rune(s)
	if len(r) > n {
		return string(r[:n]) + "..."
	}
	return s
}
