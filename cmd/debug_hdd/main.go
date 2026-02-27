package main

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"

	"github.com/shakinm/xlsReader/cfb"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	c, err := cfb.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "CFB: %v\n", err)
		return
	}

	dirs := c.GetDirs()
	var root, wordDocDir, tableDir *cfb.Directory
	for _, d := range dirs {
		name := d.Name()
		if name == "Root Entry" { root = d }
		if name == "WordDocument" { wordDocDir = d }
		if name == "1Table" || name == "0Table" {
			if tableDir == nil { tableDir = d }
		}
	}

	wdReader, _ := c.OpenObject(wordDocDir, root)
	wordDocData, _ := io.ReadAll(wdReader)
	tReader, _ := c.OpenObject(tableDir, root)
	tableData, _ := io.ReadAll(tReader)

	// FIB fields
	ccpText := binary.LittleEndian.Uint32(wordDocData[0x4C:0x50])
	ccpFtn := binary.LittleEndian.Uint32(wordDocData[0x50:0x54])
	ccpHdd := binary.LittleEndian.Uint32(wordDocData[0x54:0x58])
	ccpAtn := binary.LittleEndian.Uint32(wordDocData[0x5C:0x60])

	fmt.Printf("ccpText=%d ccpFtn=%d ccpHdd=%d ccpAtn=%d\n", ccpText, ccpFtn, ccpHdd, ccpAtn)

	fcPlcfHdd := binary.LittleEndian.Uint32(wordDocData[0xF2:0xF6])
	lcbPlcfHdd := binary.LittleEndian.Uint32(wordDocData[0xF6:0xFA])
	fmt.Printf("PlcfHdd: fc=%d lcb=%d\n", fcPlcfHdd, lcbPlcfHdd)

	plcData := tableData[fcPlcfHdd : fcPlcfHdd+lcbPlcfHdd]
	nCPs := lcbPlcfHdd / 4
	fmt.Printf("Number of CPs: %d\n", nCPs)

	cps := make([]uint32, nCPs)
	for i := uint32(0); i < nCPs; i++ {
		cps[i] = binary.LittleEndian.Uint32(plcData[i*4:])
	}
	fmt.Printf("CPs: %v\n", cps)

	// Get the full text to extract header stories
	d2, _ := cfb.OpenReader(f)
	_ = d2

	// Re-read the document to get full text
	f.Seek(0, 0)
	doc2, err := cfb.OpenReader(f)
	if err != nil {
		return
	}

	// Get WordDocument stream again for piece table
	dirs2 := doc2.GetDirs()
	var root2, wd2 *cfb.Directory
	for _, d := range dirs2 {
		if d.Name() == "Root Entry" { root2 = d }
		if d.Name() == "WordDocument" { wd2 = d }
	}
	wdR2, _ := doc2.OpenObject(wd2, root2)
	_ = wdR2

	// Use the doc package to get the full text
	f.Seek(0, 0)
	import_doc, err2 := importDoc(f)
	if err2 != nil {
		fmt.Printf("Import error: %v\n", err2)
	}
	_ = import_doc

	// Let's just manually check the header text area
	hddStart := ccpText + ccpFtn
	fmt.Printf("Header text starts at CP %d\n", hddStart)

	// Show each story
	storyNames := []string{"even header", "odd header", "even footer", "odd footer", "first header", "first footer"}
	for i := 0; i+1 < int(nCPs); i++ {
		cpStart := cps[i]
		cpEnd := cps[i+1]
		storyName := "?"
		if i%6 < len(storyNames) {
			storyName = storyNames[i%6]
		}
		fmt.Printf("Story[%d] (%s): cp[%d-%d] (absolute: %d-%d)\n",
			i, storyName, cpStart, cpEnd, hddStart+cpStart, hddStart+cpEnd)
	}
}

func importDoc(f *os.File) (interface{}, error) {
	return nil, nil
}
