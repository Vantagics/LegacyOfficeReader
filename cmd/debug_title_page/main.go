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
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	// Detailed analysis of first 40 paragraphs (title page area)
	for i := 0; i < 40 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}

		// Show hex of text
		hexStr := ""
		for _, b := range []byte(text) {
			hexStr += fmt.Sprintf("%02X ", b)
		}
		if len(hexStr) > 120 {
			hexStr = hexStr[:120] + "..."
		}

		display := strings.Map(func(r rune) rune {
			if r < 0x20 && r != '\t' {
				return '·'
			}
			return r
		}, text)

		fmt.Printf("P%d: text=%q (len=%d runes=%d)\n", i, display, len(text), len([]rune(text)))
		if hexStr != "" {
			fmt.Printf("  hex: %s\n", hexStr)
		}
		fmt.Printf("  align=%d drawn=%v runs=%d\n", p.Props.Alignment, p.DrawnImages, len(p.Runs))
		for j, r := range p.Runs {
			rText := strings.Map(func(c rune) rune {
				if c < 0x20 {
					return '·'
				}
				return c
			}, r.Text)
			fmt.Printf("  R%d: text=%q font=%q sz=%d bold=%v imgRef=%d picLoc=%d hasPicLoc=%v\n",
				j, rText, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.ImageRef, r.Props.PicLocation, r.Props.HasPicLocation)
		}
	}

	// Also show raw text first 500 chars
	rawText := d.GetText()
	runes := []rune(rawText)
	if len(runes) > 500 {
		runes = runes[:500]
	}
	fmt.Printf("\n--- Raw text first 500 runes ---\n")
	for i, r := range runes {
		if r < 0x20 {
			fmt.Printf("[%d:%02X]", i, r)
		} else {
			fmt.Printf("%c", r)
		}
	}
	fmt.Println()
}
