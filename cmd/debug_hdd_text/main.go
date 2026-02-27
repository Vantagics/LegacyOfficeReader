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

	// Get raw text and show the header area
	rawText := d.GetText()
	runes := []rune(rawText)
	fmt.Printf("Total runes: %d\n", len(runes))

	// Header starts at CP 10327 (ccpText=10327, ccpFtn=0)
	hddStart := 10327
	hddEnd := hddStart + 112 // ccpHdd=112
	if hddEnd > len(runes) {
		hddEnd = len(runes)
	}

	fmt.Printf("Header text area (CP %d-%d):\n", hddStart, hddEnd)
	for i := hddStart; i < hddEnd; i++ {
		r := runes[i]
		if r < 0x20 {
			fmt.Printf("[%d:%02X]", i-hddStart, r)
		} else {
			fmt.Printf("%c", r)
		}
	}
	fmt.Println()

	// Also show the area just before header start
	fmt.Printf("\nText before header (CP %d-%d):\n", hddStart-5, hddStart)
	for i := hddStart - 5; i < hddStart; i++ {
		r := runes[i]
		if r < 0x20 {
			fmt.Printf("[%02X]", r)
		} else {
			fmt.Printf("%c", r)
		}
	}
	fmt.Println()
}
