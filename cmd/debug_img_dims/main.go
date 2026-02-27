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
	for _, idx := range []int{0, 6, 7} {
		if idx >= len(images) { continue }
		img := images[idx]
		data := img.Data
		if img.Format == 4 { // PNG
			if len(data) > 24 && data[0] == 0x89 && data[1] == 'P' {
				w := int(data[16])<<24 | int(data[17])<<16 | int(data[18])<<8 | int(data[19])
				h := int(data[20])<<24 | int(data[21])<<16 | int(data[22])<<8 | int(data[23])
				fmt.Printf("BSE[%d]: %dx%d pixels, %d bytes\n", idx, w, h, len(data))
			}
		}
	}
}
