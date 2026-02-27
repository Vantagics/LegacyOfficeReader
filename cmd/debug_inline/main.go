package main

import (
	"encoding/binary"
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/doc"
)

func main() {
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	images := d.GetImages()
	fmt.Printf("Total images extracted: %d\n", len(images))
	for i, img := range images {
		fmt.Printf("  Image %d: format=%d size=%d bytes\n", i, img.Format, len(img.Data))
	}

	fc := d.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		return
	}

	// Find paragraphs with 0x01 in their runs
	fmt.Println("\n=== Paragraphs with inline images (0x01 in runs) ===")
	inlineCount := 0
	for i, p := range fc.Paragraphs {
		for _, r := range p.Runs {
			if strings.Contains(r.Text, "\x01") {
				fmt.Printf("  P%d: %q\n", i, truncate(r.Text, 80))
				inlineCount++
			}
		}
	}
	fmt.Printf("Total paragraphs with inline images: %d\n", inlineCount)

	// Find paragraphs with drawn images
	fmt.Println("\n=== Paragraphs with drawn object images ===")
	for i, p := range fc.Paragraphs {
		if len(p.DrawnImages) > 0 {
			text := ""
			for _, r := range p.Runs {
				text += r.Text
			}
			fmt.Printf("  P%d: DrawnImages=%v text=%q\n", i, p.DrawnImages, truncate(text, 80))
		}
	}

	// Also check the raw text for 0x01 and 0x08 positions
	text := d.GetText()
	runes := []rune(text)
	fmt.Printf("\n=== Raw text special chars ===\n")
	fmt.Printf("Total runes: %d\n", len(runes))
	for i, r := range runes {
		if r == 0x01 {
			fmt.Printf("  0x01 at CP %d\n", i)
		}
		if r == 0x08 {
			fmt.Printf("  0x08 at CP %d\n", i)
		}
	}

	// Now check the Data stream for inline image structures
	// Inline images in DOC use sprmCPicLocation to point to Data stream offset
	// where there's a PICFAndOfficeArtData structure
	adaptor, _ := cfb.OpenFile("testfie/test.doc")
	defer adaptor.CloseFile()

	var root, dataDir *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "Data":
			dataDir = dir
		}
	}

	if dataDir != nil {
		dReader, _ := adaptor.OpenObject(dataDir, root)
		dSize := binary.LittleEndian.Uint32(dataDir.StreamSize[:])
		dData := make([]byte, dSize)
		dReader.Read(dData)

		// The SpContainers we found are at offsets 4153, 111279, 457690
		// These contain blips. Let's check what's at offset 0 of the Data stream
		// and look for PICFAndOfficeArtData structures
		fmt.Printf("\n=== Data stream analysis ===\n")
		fmt.Printf("Data stream size: %d\n", len(dData))

		// Check for SpContainer records and their blip types
		for i := 0; i+8 <= len(dData); i++ {
			verInst := binary.LittleEndian.Uint16(dData[i : i+2])
			recVer := verInst & 0x0F
			recType := binary.LittleEndian.Uint16(dData[i+2 : i+4])
			recLen := binary.LittleEndian.Uint32(dData[i+4 : i+8])

			if recType == 0xF004 && recVer == 0xF && recLen > 0 && uint32(i)+8+recLen <= uint32(len(dData)) {
				// Find blip inside this SpContainer
				findBlipInContainer(dData, uint32(i)+8, uint32(i)+8+recLen, i)
				i += int(7 + recLen)
			}
		}
	}
}

func findBlipInContainer(data []byte, offset, limit uint32, containerOffset int) {
	for offset+8 <= limit {
		recType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
		recLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])
		childEnd := offset + 8 + recLen
		if childEnd > limit {
			break
		}
		if recType >= 0xF01A && recType <= 0xF02A {
			fmt.Printf("  SpContainer@%d has blip type=0x%04X size=%d\n", containerOffset, recType, recLen)
		}
		offset = childEnd
	}
}

func truncate(s string, n int) string {
	runes := []rune(s)
	if len(runes) > n {
		return string(runes[:n]) + "..."
	}
	return s
}
