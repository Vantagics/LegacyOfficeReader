package main

import (
	"fmt"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		return
	}

	// Get the full text including headers
	// We need to access the raw full text - but GetText() only returns main text
	// Let's check what the header extraction sees
	fc := d.GetFormattedContent()
	fmt.Printf("Headers: %v\n", fc.Headers)
	fmt.Printf("Footers: %v\n", fc.Footers)

	// The issue is that the header text is only "奇安" (2 chars)
	// But the full header should be something like "奇安信天眼威胁监测与分析系统"
	// Let's check if the header text contains field codes
	for i, h := range fc.Headers {
		fmt.Printf("Header[%d]: ", i)
		for _, r := range h {
			fmt.Printf("U+%04X(%c) ", r, r)
		}
		fmt.Println()
	}
	for i, f := range fc.Footers {
		fmt.Printf("Footer[%d]: ", i)
		for _, r := range f {
			if r < 0x20 {
				fmt.Printf("[%02X] ", r)
			} else {
				fmt.Printf("U+%04X(%c) ", r, r)
			}
		}
		fmt.Println()
	}
}
