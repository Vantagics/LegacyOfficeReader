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
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	masters := p.GetMasters()
	slides := p.GetSlides()

	// Count which masters are used
	masterUsage := make(map[uint32]int)
	for _, s := range slides {
		masterUsage[s.GetMasterRef()]++
	}

	fmt.Printf("Masters: %d\n", len(masters))
	for ref, m := range masters {
		fmt.Printf("\n=== Master ref=%d (used by %d slides) ===\n", ref, masterUsage[ref])
		fmt.Printf("BG: has=%v color=%s imgIdx=%d\n", m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)
		fmt.Printf("ColorScheme: %v\n", m.ColorScheme)
		fmt.Printf("DefaultTextStyles:\n")
		for i, ts := range m.DefaultTextStyles {
			fmt.Printf("  [%d] sz=%d font=%q bold=%v italic=%v color=%s\n",
				i, ts.FontSize, ts.FontName, ts.Bold, ts.Italic, ts.Color)
		}
		fmt.Printf("Shapes: %d\n", len(m.Shapes))
		for si, sh := range m.Shapes {
			desc := fmt.Sprintf("  [%d] type=%d (%dx%d @ %d,%d) img=%v(idx=%d) text=%v",
				si, sh.ShapeType, sh.Width, sh.Height, sh.Left, sh.Top, sh.IsImage, sh.ImageIdx, sh.IsText)
			if sh.FillColor != "" {
				desc += fmt.Sprintf(" fill=%s", sh.FillColor)
			}
			if sh.NoFill {
				desc += " noFill"
			}
			fmt.Println(desc)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					text := run.Text
					if len(text) > 80 {
						text = text[:80] + "..."
					}
					text = strings.ReplaceAll(text, "\n", "\\n")
					text = strings.ReplaceAll(text, "\r", "\\r")
					text = strings.ReplaceAll(text, "\x0b", "\\v")
					fmt.Printf("    P%d.R%d: font=%q sz=%d color=%s text=%q\n",
						pi, ri, run.FontName, run.FontSize, run.Color, text)
				}
			}
		}
	}
}
