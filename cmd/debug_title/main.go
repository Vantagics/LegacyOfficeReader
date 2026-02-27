package main

import (
	"fmt"
	"os"

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
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Show paragraphs 0-50 with full detail
	fmt.Printf("=== Paragraphs 0-50 Detail ===\n")
	for i := 0; i < 50 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}

		fmt.Printf("\nP%d: align=%d inTable=%v heading=%d\n", i, p.Props.Alignment, p.InTable, p.HeadingLevel)
		for ri, r := range p.Runs {
			rText := r.Text
			// Show hex for special chars
			hasSpecial := false
			for _, c := range rText {
				if c < 0x20 && c != '\t' {
					hasSpecial = true
					break
				}
			}
			if hasSpecial {
				fmt.Printf("  Run%d: %q (hex)", ri, rText)
			} else {
				fmt.Printf("  Run%d: %q", ri, rText)
			}
			fmt.Printf(" font=%q sz=%d bold=%v italic=%v color=%q",
				r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Italic, r.Props.Color)
			if r.ImageRef >= 0 {
				fmt.Printf(" imgRef=%d", r.ImageRef)
			}
			fmt.Println()
		}
		if len(p.DrawnImages) > 0 {
			fmt.Printf("  DrawnImages: %v\n", p.DrawnImages)
		}
	}

	// Also show raw text around the title area
	rawText := d.GetText()
	runes := []rune(rawText)
	fmt.Printf("\n=== Raw text first 200 chars (hex) ===\n")
	for i := 0; i < 200 && i < len(runes); i++ {
		r := runes[i]
		if r >= 0x20 && r < 0x7F {
			fmt.Printf("%c", r)
		} else if r >= 0x80 {
			fmt.Printf("%c", r)
		} else {
			fmt.Printf("[%02X]", r)
		}
	}
	fmt.Println()
}
