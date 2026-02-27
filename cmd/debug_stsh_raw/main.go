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
		os.Exit(1)
	}

	dirs := c.GetDirs()
	var root *cfb.Directory
	var wordDocDir, tableDir *cfb.Directory
	for _, d := range dirs {
		name := d.Name()
		if name == "Root Entry" {
			root = d
		}
		if name == "WordDocument" {
			wordDocDir = d
		}
		if name == "1Table" || name == "0Table" {
			if tableDir == nil {
				tableDir = d
			}
		}
	}

	if root == nil || wordDocDir == nil || tableDir == nil {
		fmt.Println("Missing streams")
		return
	}

	// Read WordDocument
	wdReader, err := c.OpenObject(wordDocDir, root)
	if err != nil {
		fmt.Fprintf(os.Stderr, "WD: %v\n", err)
		return
	}
	wordDocData, _ := io.ReadAll(wdReader)

	// Read Table
	tReader, err := c.OpenObject(tableDir, root)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Table: %v\n", err)
		return
	}
	tableData, _ := io.ReadAll(tReader)

	// Get STSH offset from FIB
	fcStshf := binary.LittleEndian.Uint32(wordDocData[0xA2:0xA6])
	lcbStshf := binary.LittleEndian.Uint32(wordDocData[0xA6:0xAA])
	fmt.Printf("STSH: fc=%d lcb=%d\n", fcStshf, lcbStshf)

	stshData := tableData[fcStshf : fcStshf+lcbStshf]

	// Parse header
	cbStshi := binary.LittleEndian.Uint16(stshData[0:2])
	cstd := binary.LittleEndian.Uint16(stshData[2:4])
	cbSTDBase := binary.LittleEndian.Uint16(stshData[4:6])
	fmt.Printf("cbStshi=%d cstd=%d cbSTDBase=%d\n", cbStshi, cstd, cbSTDBase)

	// Dump Stshi header
	fmt.Printf("Stshi header bytes: ")
	end := int(cbStshi) + 2
	if end > len(stshData) {
		end = len(stshData)
	}
	for i := 0; i < end && i < 30; i++ {
		fmt.Printf("%02X ", stshData[i])
	}
	fmt.Println()

	// Parse first few STDs
	pos := int(cbStshi) + 2
	for i := 0; i < 5 && i < int(cstd); i++ {
		if pos+2 > len(stshData) {
			break
		}
		cbStd := binary.LittleEndian.Uint16(stshData[pos:])
		pos += 2
		if cbStd == 0 {
			fmt.Printf("Style[%d]: empty\n", i)
			continue
		}
		stdData := stshData[pos : pos+int(cbStd)]
		pos += int(cbStd)

		fmt.Printf("\nStyle[%d]: cbStd=%d\n", i, cbStd)
		// Dump first 40 bytes
		dumpEnd := 40
		if dumpEnd > len(stdData) {
			dumpEnd = len(stdData)
		}
		fmt.Printf("  raw: ")
		for j := 0; j < dumpEnd; j++ {
			fmt.Printf("%02X ", stdData[j])
		}
		fmt.Println()

		// Parse fields
		word0 := binary.LittleEndian.Uint16(stdData[0:2])
		sti := word0 & 0x0FFF
		stk := (word0 >> 12) & 0x0F
		fmt.Printf("  word0=0x%04X sti=%d stk=%d\n", word0, sti, stk)

		if len(stdData) >= 4 {
			word1 := binary.LittleEndian.Uint16(stdData[2:4])
			fmt.Printf("  word1=0x%04X (istdBase=%d)\n", word1, word1)
		}
		if len(stdData) >= 6 {
			word2 := binary.LittleEndian.Uint16(stdData[4:6])
			fmt.Printf("  word2=0x%04X\n", word2)
		}
		if len(stdData) >= 8 {
			word3 := binary.LittleEndian.Uint16(stdData[6:8])
			fmt.Printf("  word3=0x%04X\n", word3)
		}

		// Name at cbSTDBase offset
		nameOff := int(cbSTDBase)
		if nameOff+2 <= len(stdData) {
			nameLen := binary.LittleEndian.Uint16(stdData[nameOff:])
			fmt.Printf("  nameOffset=%d nameLen=%d\n", nameOff, nameLen)
		}
	}
}
