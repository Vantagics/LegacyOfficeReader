package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	fmt.Printf("Headers (%d):\n", len(fc.Headers))
	for i, h := range fc.Headers {
		fmt.Printf("  [%d]: %q (len=%d)\n", i, h, len(h))
		// Show hex
		for j, b := range []byte(h) {
			if j > 60 { fmt.Print("..."); break }
			fmt.Printf("%02X ", b)
		}
		fmt.Println()
	}

	fmt.Printf("\nFooters (%d):\n", len(fc.Footers))
	for i, ft := range fc.Footers {
		fmt.Printf("  [%d]: %q (len=%d)\n", i, ft, len(ft))
	}
}
