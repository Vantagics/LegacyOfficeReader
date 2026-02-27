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
	slide := slides[70] // slide 71
	masterRef := slide.GetMasterRef()
	defaults := slide.GetDefaultTextStyles()

	fmt.Printf("Slide 71 master ref: %d\n", masterRef)
	fmt.Printf("Default text styles:\n")
	for i, d := range defaults {
		fmt.Printf("  Level %d: fontSize=%d fontName=%q bold=%v italic=%v color=%s colorRaw=0x%08X\n",
			i, d.FontSize, d.FontName, d.Bold, d.Italic, d.Color, d.ColorRaw)
	}

	// Also check what the master has
	masters := p.GetMasters()
	if m, ok := masters[masterRef]; ok {
		fmt.Printf("\nMaster %d:\n", masterRef)
		fmt.Printf("  ColorScheme: %v\n", m.ColorScheme)
		fmt.Printf("  DefaultTextStyles:\n")
		for i, d := range m.DefaultTextStyles {
			fmt.Printf("    Level %d: fontSize=%d fontName=%q bold=%v color=%s\n",
				i, d.FontSize, d.FontName, d.Bold, d.Color)
		}
	}

	// Check shape[36] text margins
	shapes := slide.GetShapes()
	if len(shapes) > 36 {
		sh := shapes[36]
		fmt.Printf("\nShape[36] text margins: L=%d T=%d R=%d B=%d anchor=%d wrap=%d\n",
			sh.TextMarginLeft, sh.TextMarginTop, sh.TextMarginRight, sh.TextMarginBottom,
			sh.TextAnchor, sh.TextWordWrap)
		fmt.Printf("Shape[36] size: %d x %d\n", sh.Width, sh.Height)
	}
	if len(shapes) > 37 {
		sh := shapes[37]
		fmt.Printf("\nShape[37] 'Agent引流' text margins: L=%d T=%d R=%d B=%d\n",
			sh.TextMarginLeft, sh.TextMarginTop, sh.TextMarginRight, sh.TextMarginBottom)
		fmt.Printf("Shape[37] size: %d x %d\n", sh.Width, sh.Height)
	}
	if len(shapes) > 38 {
		sh := shapes[38]
		fmt.Printf("\nShape[38] '虚拟交换机引流' text margins: L=%d T=%d R=%d B=%d\n",
			sh.TextMarginLeft, sh.TextMarginTop, sh.TextMarginRight, sh.TextMarginBottom)
		fmt.Printf("Shape[38] size: %d x %d\n", sh.Width, sh.Height)
	}
}
