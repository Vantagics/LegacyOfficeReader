package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "parse: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		return
	}

	// Find CP range of paragraph 166 by walking through text
	text := d.GetText()
	runes := []rune(text)
	cpPos := uint32(0)
	paraIdx := 0
	start := 0
	var targetCPStart, targetCPEnd uint32
	for i, r := range runes {
		if r == '\r' || r == 0x07 {
			if paraIdx == 165 {
				targetCPStart = cpPos
				targetCPEnd = cpPos + uint32(i-start)
				break
			}
			cpPos = uint32(i + 1)
			start = i + 1
			paraIdx++
		}
	}

	fmt.Printf("Paragraph 166 CP range: [%d, %d] (len=%d)\n\n", targetCPStart, targetCPEnd, targetCPEnd-targetCPStart)

	// Show the actual text at this CP range
	if int(targetCPEnd) <= len(runes) {
		paraText := string(runes[targetCPStart:targetCPEnd])
		fmt.Printf("Text length: %d runes\n", len([]rune(paraText)))
		if len(paraText) > 200 {
			fmt.Printf("Text: %q...\n\n", paraText[:200])
		} else {
			fmt.Printf("Text: %q\n\n", paraText)
		}
	}
}
