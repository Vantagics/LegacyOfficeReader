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

	// Show raw text for first 100 characters with positions
	text := d.GetText()
	runes := []rune(text)
	fmt.Println("=== First 50 runes of main body text ===")
	for i := 0; i < 50 && i < len(runes); i++ {
		r := runes[i]
		if r == '\r' {
			fmt.Printf("  [%d] \\r (0x0D)\n", i)
		} else if r == '\t' {
			fmt.Printf("  [%d] \\t (0x09)\n", i)
		} else if r == 0x08 {
			fmt.Printf("  [%d] OBJ (0x08) - drawn object\n", i)
		} else if r == 0x01 {
			fmt.Printf("  [%d] IMG (0x01) - inline image\n", i)
		} else if r < 0x20 {
			fmt.Printf("  [%d] 0x%02X\n", i, r)
		} else {
			fmt.Printf("  [%d] '%c' (U+%04X)\n", i, r, r)
		}
	}

	// The 0x08 at positions 6 and 9 are drawn objects
	// Let's check what styles/formatting they have
	fc := d.GetFormattedContent()
	if fc == nil {
		return
	}

	fmt.Println("\n=== Title page paragraphs with their styles ===")
	for i := 0; i < 15 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			t := r.Text
			for _, c := range t {
				if c == 0x08 {
					text += "[OBJ]"
				} else if c == 0x01 {
					text += "[IMG]"
				} else if c == '\t' {
					text += "[TAB]"
				} else if c < 0x20 {
					text += fmt.Sprintf("[0x%02X]", c)
				} else {
					text += string(c)
				}
			}
		}
		fmt.Printf("P%d align=%d: %q\n", i, p.Props.Alignment, text)
		for j, r := range p.Runs {
			fmt.Printf("  Run%d: font=%q sz=%d bold=%v text=%q\n",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Text)
		}
	}
}
