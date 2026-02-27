package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "parse: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("no formatted content")
		return
	}

	// Find paragraphs with long text (potential duplicates)
	for i, p := range fc.Paragraphs {
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		text = strings.TrimSpace(text)

		// Check for duplicate content within the same paragraph
		if len(text) > 100 {
			half := len(text) / 2
			first := text[:half]
			second := text[half:]
			// Check if the second half starts with the same content as the first half
			minCheck := 50
			if len(first) >= minCheck && len(second) >= minCheck {
				if first[:minCheck] == second[:minCheck] {
					fmt.Printf("[%3d] DUPLICATE DETECTED! len=%d\n", i+1, len(text))
					fmt.Printf("  First 80 chars:  %s\n", text[:80])
					fmt.Printf("  Second half 80:  %s\n", second[:80])
					fmt.Printf("  Runs: %d\n", len(p.Runs))
					for j, r := range p.Runs {
						rt := r.Text
						if len(rt) > 60 {
							rt = rt[:60] + "..."
						}
						fmt.Printf("    Run[%d] font=%q sz=%d text=%q\n", j, r.Props.FontName, r.Props.FontSize, rt)
					}
					fmt.Println()
				}
			}
		}

		// Also show paragraphs containing "传感器" with their run details
		if strings.Contains(text, "传感器主要负责") {
			fmt.Printf("[%3d] Contains '传感器主要负责' len=%d runs=%d\n", i+1, len(text), len(p.Runs))
			for j, r := range p.Runs {
				rt := r.Text
				if len(rt) > 80 {
					rt = rt[:80] + "..."
				}
				fmt.Printf("    Run[%d] font=%q sz=%d text=%q\n", j, r.Props.FontName, r.Props.FontSize, rt)
			}
			fmt.Println()
		}
	}
}
