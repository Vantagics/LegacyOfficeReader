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
	for _, idx := range []int{0, 6, 7, 8, 9, 10} {
		if idx >= len(images) { continue }
		img := images[idx]
		data := img.Data
		if img.Format == 4 { // PNG
			if len(data) > 24 && data[0] == 0x89 && data[1] == 'P' {
				w := int(data[16])<<24 | int(data[17])<<16 | int(data[18])<<8 | int(data[19])
				h := int(data[20])<<24 | int(data[21])<<16 | int(data[22])<<8 | int(data[23])
				
				// Look for pHYs chunk
				dpi := 96
				for i := 8; i+16 < len(data); i++ {
					if data[i] == 'p' && data[i+1] == 'H' && data[i+2] == 'Y' && data[i+3] == 's' && i+16 <= len(data) {
						pxPerUnitX := int(data[i+4])<<24 | int(data[i+5])<<16 | int(data[i+6])<<8 | int(data[i+7])
						pxPerUnitY := int(data[i+8])<<24 | int(data[i+9])<<16 | int(data[i+10])<<8 | int(data[i+11])
						unit := data[i+12]
						if unit == 1 && pxPerUnitX > 0 {
							dpi = pxPerUnitX * 254 / 10000
						}
						fmt.Printf("BSE[%d]: %dx%d px, pHYs: %d x %d px/m, unit=%d → DPI=%d\n",
							idx, w, h, pxPerUnitX, pxPerUnitY, unit, dpi)
						break
					}
				}
				if dpi == 96 {
					fmt.Printf("BSE[%d]: %dx%d px, no pHYs chunk → default 96 DPI\n", idx, w, h)
				}
			}
		}
	}
}
