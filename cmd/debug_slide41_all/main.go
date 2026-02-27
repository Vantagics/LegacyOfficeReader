package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	sl := slides[40] // slide 41
	shapes := sl.GetShapes()
	fmt.Printf("Slide 41: %d shapes\n", len(shapes))
	for si, sh := range shapes {
		if sh.FillColor == "000000" || (len(sh.Paragraphs) > 0 && sh.FillColor != "") {
			fmt.Printf("Shape[%d] type=%d text=%v fill=%q noFill=%v fillRaw=0x%08X paras=%d pos=(%d,%d) sz=(%d,%d)\n",
				si, sh.ShapeType, sh.IsText, sh.FillColor, sh.NoFill, sh.FillColorRaw, len(sh.Paragraphs),
				sh.Left, sh.Top, sh.Width, sh.Height)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					t := run.Text
					if len(t) > 30 {
						t = t[:30] + "..."
					}
					fmt.Printf("  P[%d]R[%d] color=%q raw=0x%08X sz=%d: %q\n",
						pi, ri, run.Color, run.ColorRaw, run.FontSize, t)
				}
			}
		}
	}
	
	// Also find shapes with text containing "数据库审计" or "堡垒机"
	fmt.Println("\n--- Searching for specific text ---")
	for si, sh := range shapes {
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				if run.Text == "数据库审计系统" || run.Text == "堡垒机" || run.Text == "特权账号管理系统" || run.Text == "API监测系统" || run.Text == "跨境数据监测系统" {
					fmt.Printf("Shape[%d] type=%d fill=%q noFill=%v fillRaw=0x%08X: %q color=%q raw=0x%08X\n",
						si, sh.ShapeType, sh.FillColor, sh.NoFill, sh.FillColorRaw, run.Text, run.Color, run.ColorRaw)
				}
			}
		}
	}
}
