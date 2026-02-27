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

	rawText := d.GetText()
	
	// Search for key text
	searches := []string{"天眼", "威胁", "监测", "奇安信", "版本", "修订", "目录", "概述", "部署"}
	for _, s := range searches {
		idx := strings.Index(rawText, s)
		if idx >= 0 {
			// Show context around the match
			runes := []rune(rawText)
			runeIdx := 0
			byteIdx := 0
			for byteIdx < idx {
				byteIdx += len(string(runes[runeIdx]))
				runeIdx++
			}
			start := runeIdx - 20
			if start < 0 {
				start = 0
			}
			end := runeIdx + 40
			if end > len(runes) {
				end = len(runes)
			}
			context := string(runes[start:end])
			// Show with hex for special chars
			var display strings.Builder
			for _, r := range context {
				if r < 0x20 && r != '\t' {
					fmt.Fprintf(&display, "[%02X]", r)
				} else {
					display.WriteRune(r)
				}
			}
			fmt.Printf("Found %q at rune %d: %s\n", s, runeIdx, display.String())
		} else {
			fmt.Printf("NOT FOUND: %q\n", s)
		}
	}

	// Show paragraphs 50-80 (revision history / TOC area)
	fc := d.GetFormattedContent()
	fmt.Printf("\n=== Paragraphs 50-100 ===\n")
	for i := 50; i < 100 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.InTable {
			flags += " TABLE"
		}
		if p.TableRowEnd {
			flags += " ROWEND"
		}
		if p.IsTOC {
			flags += fmt.Sprintf(" TOC%d", p.TOCLevel)
		}
		if p.HasPageBreak {
			flags += " PB"
		}
		if p.IsSectionBreak {
			flags += " SECT"
		}
		fmt.Printf("P%d: %q%s\n", i, text, flags)
	}

	// Show paragraphs 130-170 (inline image area)
	fmt.Printf("\n=== Paragraphs 130-170 ===\n")
	for i := 130; i < 170 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.HasPageBreak {
			flags += " PB"
		}
		for _, r := range p.Runs {
			if r.ImageRef >= 0 {
				flags += fmt.Sprintf(" IMG[%d]", r.ImageRef)
			}
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN%v", p.DrawnImages)
		}
		fmt.Printf("P%d: %q%s\n", i, text, flags)
	}

	// Show paragraphs 210-220 (drawn object area)
	fmt.Printf("\n=== Paragraphs 210-225 ===\n")
	for i := 210; i < 225 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		flags := ""
		if p.HeadingLevel > 0 {
			flags += fmt.Sprintf(" H%d", p.HeadingLevel)
		}
		if p.HasPageBreak {
			flags += " PB"
		}
		if len(p.DrawnImages) > 0 {
			flags += fmt.Sprintf(" DRAWN%v", p.DrawnImages)
		}
		fmt.Printf("P%d: %q%s\n", i, text, flags)
	}
}
