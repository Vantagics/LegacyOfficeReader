package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Parse error: %v\n", err)
		os.Exit(1)
	}

	masters := p.GetMasters()
	for ref, m := range masters {
		fmt.Printf("Master ref=%d, shapes=%d, colorScheme=%v\n", ref, len(m.Shapes), m.ColorScheme)
		for i, sh := range m.Shapes {
			extra := ""
			if sh.IsImage {
				extra = fmt.Sprintf(" [IMAGE idx=%d]", sh.ImageIdx)
			}
			if sh.FillColor != "" {
				extra += fmt.Sprintf(" fill=%s", sh.FillColor)
			}
			if sh.NoFill {
				extra += " noFill"
			}
			fmt.Printf("  shape[%d]: type=%d pos=(%d,%d) size=(%d,%d) text=%v%s\n",
				i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsText, extra)
		}
		fmt.Println()
	}
}
