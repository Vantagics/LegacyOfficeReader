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

	names := d.GetStyles()
	stis := d.GetStyleSTIs()
	for i, name := range names {
		if name != "" || i < 20 {
			fmt.Printf("Style[%d]: sti=%d name=%q\n", i, stis[i], name)
		}
	}
}
