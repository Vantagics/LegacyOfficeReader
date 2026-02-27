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
	for _, idx := range []int{8, 9, 10} {
		if idx >= len(images) { continue }
		img := images[idx]
		data := img.Data
		if img.Format == 4 { // PNG
			if len(data) > 24 && data[0] == 0x89 && data[1] == 'P' {
				w := int(data[16])<<24 | int(data[17])<<16 | int(data[18])<<8 | int(data[19])
				h := int(data[20])<<24 | int(data[21])<<16 | int(data[22])<<8 | int(data[23])
				// Calculate EMU at 96 DPI
				emuW := int64(w) * 914400 / 96
				emuH := int64(h) * 914400 / 96
				maxW := int64(5274310)
				if emuW > maxW {
					scale := float64(maxW) / float64(emuW)
					emuH = int64(float64(emuH) * scale)
					emuW = maxW
				}
				fmt.Printf("BSE[%d]: %dx%d px → %dx%d EMU (%.1fx%.1f cm)\n",
					idx, w, h, emuW, emuH, float64(emuW)/360000, float64(emuH)/360000)
			}
		}
	}
}
