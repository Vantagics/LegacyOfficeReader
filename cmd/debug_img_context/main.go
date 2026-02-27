package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		return
	}

	// Show paragraphs around image placeholders
	for i, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			if strings.Contains(r.Text, "\x01") {
				// Show context: 2 paragraphs before and after
				start := i - 2
				if start < 0 {
					start = 0
				}
				end := i + 3
				if end > len(fc.Paragraphs) {
					end = len(fc.Paragraphs)
				}
				fmt.Printf("=== Image at P%d ===\n", i)
				for j := start; j < end; j++ {
					pp := fc.Paragraphs[j]
					text := ""
					for _, rr := range pp.Runs {
						t := rr.Text
						t = strings.ReplaceAll(t, "\x01", "[IMG]")
						text += t
					}
					marker := "  "
					if j == i {
						marker = ">>"
					}
					flags := ""
					if pp.HeadingLevel > 0 {
						flags += fmt.Sprintf(" H%d", pp.HeadingLevel)
					}
					if pp.HasPageBreak {
						flags += " PB"
					}
					fmt.Printf("%s P%d%s: %q\n", marker, j, flags, truncate(text, 80))
				}
				fmt.Println()
			}
		}
	}
}

func truncate(s string, maxLen int) string {
	runes := []rune(s)
	if len(runes) <= maxLen {
		return s
	}
	return string(runes[:maxLen]) + "..."
}
