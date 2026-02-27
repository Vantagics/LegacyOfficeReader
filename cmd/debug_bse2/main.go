package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	images := d.GetImages()
	fmt.Printf("Total images: %d\n", len(images))
	for i, img := range images {
		ext := "?"
		switch img.Format {
		case 0:
			ext = "EMF"
		case 4:
			ext = "PNG"
		case 3:
			ext = "JPEG"
		}
		// Check first bytes for format verification
		header := ""
		if len(img.Data) >= 4 {
			header = fmt.Sprintf("%02X%02X%02X%02X", img.Data[0], img.Data[1], img.Data[2], img.Data[3])
		}
		fmt.Printf("  BSE[%d]: %s %d bytes header=%s\n", i, ext, len(img.Data), header)
	}

	// The shapes in Data stream have:
	// SPID=1027 pib=3 → BSE[2] (PNG 5KB)
	// SPID=1026 pib=2 → BSE[1] (EMF 346KB)
	// SPID=1025 pib=1 → BSE[0] (PNG 110KB)
	//
	// PlcSpaMom:
	// CP=12 SPID=2050 → Data SPID=1026 → pib=2 → BSE[1] (EMF 346KB)
	// CP=15 SPID=2051 → Data SPID=1027 → pib=3 → BSE[2] (PNG 5KB)
	// CP=108 SPID=2049 → Data SPID=1025 → pib=1 → BSE[0] (PNG 110KB)
	//
	// Title page 0x08 at CP 6 and CP 9
	// CP 6 is before CP 12 (shape anchor), CP 9 is before CP 15
	// So likely:
	// 0x08 at CP 6 → shape anchored at CP 12 → BSE[1] (EMF 346KB) - likely the logo
	// 0x08 at CP 9 → shape anchored at CP 15 → BSE[2] (PNG 5KB) - likely product name
	//
	// The inline images (0x01) at P140, P143, P157 need to map to BSE[3-7]
	// or some subset of BSE images

	fmt.Println("\nShape-to-image mapping:")
	fmt.Println("  Title shape 1 (CP 6): BSE[1] EMF 346KB (logo?)")
	fmt.Println("  Title shape 2 (CP 9): BSE[2] PNG 5KB (product name?)")
	fmt.Println("  Shape 3 (CP 108): BSE[0] PNG 110KB")
}
