package main

import (
	"fmt"
	"os"

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

	// Show paragraph 166 (0-indexed: 165) in detail
	idx := 165
	p := fc.Paragraphs[idx]
	fmt.Printf("Paragraph %d: runs=%d heading=%d\n\n", idx+1, len(p.Runs), p.HeadingLevel)

	totalText := ""
	for i, r := range p.Runs {
		fmt.Printf("Run[%d]:\n", i)
		fmt.Printf("  Font: %q\n", r.Props.FontName)
		fmt.Printf("  Size: %d (half-points)\n", r.Props.FontSize)
		fmt.Printf("  Bold: %v, Italic: %v\n", r.Props.Bold, r.Props.Italic)
		fmt.Printf("  Text len: %d chars\n", len([]rune(r.Text)))
		if len(r.Text) > 200 {
			fmt.Printf("  Text: %q...\n", r.Text[:200])
		} else {
			fmt.Printf("  Text: %q\n", r.Text)
		}
		totalText += r.Text
		fmt.Println()
	}

	fmt.Printf("Total text length: %d chars\n", len([]rune(totalText)))

	// Check if text is duplicated
	runes := []rune(totalText)
	if len(runes) > 200 {
		half := len(runes) / 2
		first50 := string(runes[:50])
		second50 := string(runes[half : half+50])
		fmt.Printf("\nFirst 50 chars:  %q\n", first50)
		fmt.Printf("Chars at half:   %q\n", second50)
		if first50 == second50 {
			fmt.Println("*** TEXT IS DUPLICATED ***")
		}
	}
}
