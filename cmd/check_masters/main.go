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
	masters := p.GetMasters()

	// Count slides per master
	masterCount := map[uint32]int{}
	for _, s := range slides {
		masterCount[s.GetMasterRef()]++
	}

	fmt.Printf("Total slides: %d\n", len(slides))
	fmt.Printf("Total masters parsed: %d\n", len(masters))
	fmt.Printf("\nSlides per master ref:\n")
	for ref, count := range masterCount {
		_, found := masters[ref]
		fmt.Printf("  ref=%d: %d slides, master found=%v\n", ref, count, found)
	}

	fmt.Printf("\nAll parsed master refs:\n")
	for ref := range masters {
		fmt.Printf("  ref=%d\n", ref)
	}
}
