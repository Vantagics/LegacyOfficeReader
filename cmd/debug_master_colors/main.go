package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		panic(err)
	}

	masters := p.GetMasters()
	for ref, m := range masters {
		fmt.Printf("Master ref=%d scheme=%v\n", ref, m.ColorScheme)
		for i, style := range m.DefaultTextStyles {
			if style.Color != "" || style.ColorRaw != 0 {
				fmt.Printf("  Level %d: color=%s raw=0x%08X\n", i, style.Color, style.ColorRaw)
			}
		}
		// Check master shape text colors
		for si, sh := range m.Shapes {
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					if run.ColorRaw != 0 {
						flag := run.ColorRaw >> 24
						text := run.Text
						if len(text) > 20 {
							text = text[:20]
						}
						fmt.Printf("  Shape %d P%d R%d: color=%s raw=0x%08X flag=0x%02X text=%q\n",
							si, pi, ri, run.Color, run.ColorRaw, flag, text)
					}
				}
			}
		}
	}
}
