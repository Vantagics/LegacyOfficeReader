package main

import (
	"fmt"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	fc := d.GetFormattedContent()
	paras := fc.Paragraphs
	images := d.GetImages()

	fmt.Printf("Total paragraphs: %d\n", len(paras))
	fmt.Printf("Total images: %d\n", len(images))

	// Show first 20 paragraphs in detail
	for i := 0; i < 20 && i < len(paras); i++ {
		p := paras[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		// Clean for display
		displayText := ""
		for _, c := range text {
			if c == '\t' {
				displayText += "\\t"
			} else if c == '\x08' {
				displayText += "\\x08"
			} else if c == '\x01' {
				displayText += "\\x01"
			} else if c < 0x20 && c != '\n' {
				displayText += fmt.Sprintf("\\x%02x", c)
			} else {
				displayText += string(c)
			}
		}
		fmt.Printf("\nP[%d]: text=%q\n", i, displayText)
		fmt.Printf("  Heading=%d, InTable=%v, IsSectionBreak=%v, SectionType=%d\n",
			p.HeadingLevel, p.InTable, p.IsSectionBreak, p.SectionType)
		fmt.Printf("  PageBreakBefore=%v, HasPageBreak=%v\n", p.PageBreakBefore, p.HasPageBreak)
		fmt.Printf("  DrawnImages=%v, TextBoxText=%q\n", p.DrawnImages, p.TextBoxText)
		fmt.Printf("  Alignment=%d, AlignmentSet=%v\n", p.Props.Alignment, p.Props.AlignmentSet)
		for ri, r := range p.Runs {
			fmt.Printf("  Run[%d]: text=%q, ImageRef=%d, HasPicLoc=%v, PicLoc=%d\n",
				ri, r.Text, r.ImageRef, r.Props.HasPicLocation, r.Props.PicLocation)
			fmt.Printf("    Font=%q, Size=%d, Bold=%v\n", r.Props.FontName, r.Props.FontSize, r.Props.Bold)
		}
	}

	// Show image sizes
	fmt.Println("\n=== Image Sizes ===")
	for i, img := range images {
		fmt.Printf("BSE[%d]: %d bytes, format=%d\n", i, len(img.Data), img.Format)
	}
}
