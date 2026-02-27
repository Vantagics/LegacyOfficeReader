package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Parse error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	// Count which slides use master 2147483734 (the one with the tiger watermark)
	tigerMaster := uint32(2147483734)
	fmt.Printf("Total slides: %d\n", len(slides))
	fmt.Printf("Slides using master %d (tiger watermark):\n", tigerMaster)
	count := 0
	for i, s := range slides {
		if s.GetMasterRef() == tigerMaster {
			count++
			fmt.Printf("  slide %d\n", i+1)
		}
	}
	fmt.Printf("Total: %d slides with tiger watermark\n", count)

	// Also show all unique master refs and their slide counts
	fmt.Println("\nAll master refs:")
	refCount := make(map[uint32]int)
	for _, s := range slides {
		refCount[s.GetMasterRef()]++
	}
	for ref, cnt := range refCount {
		fmt.Printf("  ref=%d: %d slides\n", ref, cnt)
	}
}
