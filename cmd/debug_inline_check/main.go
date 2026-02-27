package main

import (
	"fmt"
	"strings"
	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}
	
	for i, p := range fc.Paragraphs {
		for j, r := range p.Runs {
			if strings.Contains(r.Text, "\x01") {
				text := ""
				for _, rr := range p.Runs {
					text += rr.Text
				}
				fmt.Printf("P[%d] Run[%d] has inline image (0x01), ImageRef=%d, PicLoc=%d, HasPicLoc=%v\n",
					i, j, r.ImageRef, r.Props.PicLocation, r.Props.HasPicLocation)
				fmt.Printf("  Full text: %q\n", truncate(text, 80))
			}
		}
	}
}

func truncate(s string, n int) string {
	runes := []rune(s)
	if len(runes) > n {
		return string(runes[:n]) + "..."
	}
	return s
}
