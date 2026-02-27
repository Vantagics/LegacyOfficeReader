package main

import (
	"fmt"
	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	images := d.GetImages()
	fmt.Printf("Total images: %d\n", len(images))
	for i, img := range images {
		fmt.Printf("  BSE[%d]: ext=%q size=%d bytes\n", i, img.Extension(), len(img.Data))
	}
}
