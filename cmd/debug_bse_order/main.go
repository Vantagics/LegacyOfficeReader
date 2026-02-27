package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/common"
	"github.com/shakinm/xlsReader/doc"
)

func main() {
	adaptor, err := cfb.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer adaptor.CloseFile()

	var wordDoc, root *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "WordDocument":
			wordDoc = dir
		case "Root Entry":
			root = dir
		}
	}

	wordDocReader, _ := adaptor.OpenObject(wordDoc, root)
	wordDocSize := binary.LittleEndian.Uint32(wordDoc.StreamSize[:])
	wordDocData := make([]byte, wordDocSize)
	wordDocReader.Read(wordDocData)

	// Open the document to get images
	d, err := doc.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	images := d.GetImages()
	fmt.Printf("Total images: %d\n\n", len(images))
	for i, img := range images {
		ext := (&common.Image{Format: img.Format}).Extension()
		fmt.Printf("BSE[%d]: format=%d, ext=%s, size=%d bytes\n", i, img.Format, ext, len(img.Data))

		// Try to identify the image content
		if img.Format == common.ImageFormatPNG && len(img.Data) > 24 {
			w := int(img.Data[16])<<24 | int(img.Data[17])<<16 | int(img.Data[18])<<8 | int(img.Data[19])
			h := int(img.Data[20])<<24 | int(img.Data[21])<<16 | int(img.Data[22])<<8 | int(img.Data[23])
			fmt.Printf("  PNG dimensions: %dx%d\n", w, h)
		}
		if img.Format == common.ImageFormatEMF && len(img.Data) >= 40 {
			left := int32(img.Data[24]) | int32(img.Data[25])<<8 | int32(img.Data[26])<<16 | int32(img.Data[27])<<24
			top := int32(img.Data[28]) | int32(img.Data[29])<<8 | int32(img.Data[30])<<16 | int32(img.Data[31])<<24
			right := int32(img.Data[32]) | int32(img.Data[33])<<8 | int32(img.Data[34])<<16 | int32(img.Data[35])<<24
			bottom := int32(img.Data[36]) | int32(img.Data[37])<<8 | int32(img.Data[38])<<16 | int32(img.Data[39])<<24
			fmt.Printf("  EMF frame: (%d,%d)-(%d,%d) = %dx%d (0.01mm)\n", left, top, right, bottom, right-left, bottom-top)
		}
	}
}
