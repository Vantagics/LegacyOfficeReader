package main

import (
	"encoding/binary"
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
	for i, img := range images {
		fmt.Printf("\n=== Image %d: format=%d size=%d ===\n", i, img.Format, len(img.Data))
		data := img.Data
		
		// Show first 100 bytes hex
		n := 100
		if n > len(data) {
			n = len(data)
		}
		fmt.Printf("First %d bytes: ", n)
		for j := 0; j < n; j++ {
			fmt.Printf("%02X ", data[j])
			if (j+1)%16 == 0 {
				fmt.Println()
				fmt.Print("  ")
			}
		}
		fmt.Println()

		// For EMF files, parse the header
		if img.Format == 0 && len(data) >= 88 { // EMF
			fmt.Println("EMF Header:")
			recType := binary.LittleEndian.Uint32(data[0:4])
			recSize := binary.LittleEndian.Uint32(data[4:8])
			fmt.Printf("  RecordType: 0x%08X, RecordSize: %d\n", recType, recSize)
			
			// Bounds rectangle (offset 8-23)
			boundsLeft := int32(binary.LittleEndian.Uint32(data[8:12]))
			boundsTop := int32(binary.LittleEndian.Uint32(data[12:16]))
			boundsRight := int32(binary.LittleEndian.Uint32(data[16:20]))
			boundsBottom := int32(binary.LittleEndian.Uint32(data[20:24]))
			fmt.Printf("  Bounds: left=%d top=%d right=%d bottom=%d\n", boundsLeft, boundsTop, boundsRight, boundsBottom)
			fmt.Printf("  Bounds size: %d x %d (device units)\n", boundsRight-boundsLeft, boundsBottom-boundsTop)
			
			// Frame rectangle (offset 24-39) - in 0.01mm units
			if len(data) >= 40 {
				frameLeft := int32(binary.LittleEndian.Uint32(data[24:28]))
				frameTop := int32(binary.LittleEndian.Uint32(data[28:32]))
				frameRight := int32(binary.LittleEndian.Uint32(data[32:36]))
				frameBottom := int32(binary.LittleEndian.Uint32(data[36:40]))
				fmt.Printf("  Frame: left=%d top=%d right=%d bottom=%d (0.01mm)\n", frameLeft, frameTop, frameRight, frameBottom)
				fw := frameRight - frameLeft
				fh := frameBottom - frameTop
				fmt.Printf("  Frame size: %d x %d (0.01mm) = %.1f x %.1f mm\n", fw, fh, float64(fw)/100, float64(fh)/100)
				// Convert to EMU: 1mm = 36000 EMU, so 0.01mm = 360 EMU
				emuW := int64(fw) * 360
				emuH := int64(fh) * 360
				fmt.Printf("  Frame size: %d x %d EMU\n", emuW, emuH)
			}
			
			// Signature (offset 40-43)
			if len(data) >= 44 {
				sig := binary.LittleEndian.Uint32(data[40:44])
				fmt.Printf("  Signature: 0x%08X\n", sig)
			}
			
			// Version (offset 44-47)
			if len(data) >= 48 {
				ver := binary.LittleEndian.Uint32(data[44:48])
				fmt.Printf("  Version: 0x%08X\n", ver)
			}
			
			// Size (offset 48-51)
			if len(data) >= 52 {
				sz := binary.LittleEndian.Uint32(data[48:52])
				fmt.Printf("  Bytes: %d\n", sz)
			}
			
			// Device size (offset 68-75)
			if len(data) >= 76 {
				devW := binary.LittleEndian.Uint32(data[68:72])
				devH := binary.LittleEndian.Uint32(data[72:76])
				fmt.Printf("  Device: %d x %d pixels\n", devW, devH)
			}
			
			// Millimeters size (offset 76-83)
			if len(data) >= 84 {
				mmW := binary.LittleEndian.Uint32(data[76:80])
				mmH := binary.LittleEndian.Uint32(data[80:84])
				fmt.Printf("  Millimeters: %d x %d mm\n", mmW, mmH)
			}
		}
		
		// For PNG files, parse dimensions
		if img.Format == 4 && len(data) >= 24 { // PNG
			if data[0] == 0x89 && data[1] == 'P' && data[2] == 'N' && data[3] == 'G' {
				w := int(data[16])<<24 | int(data[17])<<16 | int(data[18])<<8 | int(data[19])
				h := int(data[20])<<24 | int(data[21])<<16 | int(data[22])<<8 | int(data[23])
				fmt.Printf("PNG dimensions: %d x %d pixels\n", w, h)
			} else {
				fmt.Printf("PNG: no signature found, first bytes: %02X %02X %02X %02X\n", data[0], data[1], data[2], data[3])
			}
		}
	}
}
