package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, _ := os.Open("testfie/test.doc")
	defer f.Close()
	d, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	images := d.GetImages()
	fmt.Printf("Total BSE images: %d\n", len(images))
	for i, img := range images {
		fmt.Printf("  BSE[%d]: format=%d size=%d bytes\n", i, img.Format, len(img.Data))
	}
}
