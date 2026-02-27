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
	slide := slides[64]

	fmt.Printf("Slide 65: MasterRef=%d\n", slide.GetMasterRef())

	// Check master text styles
	masters := p.GetMasters()
	for ref, m := range masters {
		if ref == slide.GetMasterRef() {
			fmt.Printf("\nMaster %d:\n", ref)
			fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
			fmt.Printf("  TextStyles:\n")
			for level, style := range m.DefaultTextStyles {
				fmt.Printf("    Level %d: FontSize=%d Color=%q ColorRaw=0x%08X Bold=%v Font=%q\n",
					level, style.FontSize, style.Color, style.ColorRaw, style.Bold, style.FontName)
			}
			break
		}
	}

	// Also check slide's own default text styles
	styles := slide.GetDefaultTextStyles()
	fmt.Println("\nSlide DefaultTextStyles:")
	for level, style := range styles {
		fmt.Printf("  Level %d: FontSize=%d Color=%q ColorRaw=0x%08X Bold=%v Font=%q\n",
			level, style.FontSize, style.Color, style.ColorRaw, style.Bold, style.FontName)
	}
}
