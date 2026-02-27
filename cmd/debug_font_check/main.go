package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	// Check font distribution
	fontDist := make(map[string]int)
	noFontCount := 0
	totalRuns := 0

	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					totalRuns++
					if run.FontName == "" {
						noFontCount++
					} else {
						fontDist[run.FontName]++
					}
				}
			}
		}
	}

	fmt.Printf("=== Font Distribution ===\n")
	fmt.Printf("Total runs: %d\n", totalRuns)
	fmt.Printf("Runs without font: %d\n", noFontCount)
	for font, count := range fontDist {
		fmt.Printf("  %s: %d\n", font, count)
	}

	// Check text color distribution
	fmt.Printf("\n=== Text Color Distribution ===\n")
	colorDist := make(map[string]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					c := run.Color
					if c == "" {
						c = "(none)"
					}
					colorDist[c]++
				}
			}
		}
	}
	for color, count := range colorDist {
		if count > 10 {
			fmt.Printf("  %s: %d\n", color, count)
		}
	}

	// Check alignment distribution
	fmt.Printf("\n=== Alignment Distribution ===\n")
	algnDist := make(map[uint8]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				algnDist[para.Alignment]++
			}
		}
	}
	algnNames := []string{"left", "center", "right", "justify"}
	for algn, count := range algnDist {
		name := "unknown"
		if int(algn) < len(algnNames) {
			name = algnNames[algn]
		}
		fmt.Printf("  %s(%d): %d\n", name, algn, count)
	}

	// Check shapes with text but no font name (these will use default)
	fmt.Printf("\n=== Shapes with no font name (first 10) ===\n")
	count := 0
	for si, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if run.FontName == "" && strings.TrimSpace(run.Text) != "" {
						fmt.Printf("  Slide %d: \"%s\" sz=%d color=%s\n",
							si+1, truncate(run.Text, 30), run.FontSize, run.Color)
						count++
						if count >= 10 {
							goto done
						}
					}
				}
			}
		}
	}
done:

	// Check slide 41 specifically (the one with opacity issues)
	fmt.Printf("\n=== Slide 41 Detail ===\n")
	if len(slides) >= 41 {
		s := slides[40]
		for i, sh := range s.GetShapes() {
			if len(sh.Paragraphs) > 0 {
				textSnippet := ""
				for _, para := range sh.Paragraphs {
					for _, run := range para.Runs {
						t := strings.TrimSpace(run.Text)
						if t != "" {
							textSnippet = truncate(t, 30)
							break
						}
					}
					if textSnippet != "" {
						break
					}
				}
				fmt.Printf("  [%d] fill=%s opacity=%d noFill=%v text=%s\n",
					i, sh.FillColor, sh.FillOpacity, sh.NoFill, textSnippet)
			}
		}
	}
}

func truncate(s string, n int) string {
	runes := []rune(s)
	if len(runes) > n {
		return string(runes[:n]) + "..."
	}
	return s
}
