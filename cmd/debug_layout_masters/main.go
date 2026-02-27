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
	seen := make(map[uint32]bool)
	layoutIdx := 0
	for _, s := range slides {
		ref := s.GetMasterRef()
		if seen[ref] {
			continue
		}
		seen[ref] = true
		layoutIdx++
		fmt.Printf("Layout %d: masterRef=%d\n", layoutIdx, ref)
	}
}
