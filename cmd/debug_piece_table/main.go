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
	ccpMcr := binary.LittleEndian.Uint32(wordDocData[0x58:0x5C])
	ccpAtn := binary.LittleEndian.Uint32(wordDocData[0x5C:0x60])
	ccpEdn := binary.LittleEndian.Uint32(wordDocData[0x60:0x64])
	ccpTxbx := binary.LittleEndian.Uint32(wordDocData[0x64:0x68])
	ccpHdrTxbx := binary.LittleEndian.Uint32(wordDocData[0x68:0x6C])

	totalCPs := ccpText + ccpFtn + ccpHdd + ccpMcr + ccpAtn + ccpEdn + ccpTxbx + ccpHdrTxbx
	fmt.Printf("ccpText=%d ccpFtn=%d ccpHdd=%d ccpMcr=%d ccpAtn=%d ccpEdn=%d ccpTxbx=%d ccpHdrTxbx=%d\n",
		ccpText, ccpFtn, ccpHdd, ccpMcr, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx)
	fmt.Printf("Total CPs (sum): %d\n", totalCPs)
	fmt.Printf("Total CPs (with final \\r): %d\n", totalCPs+1)

	// Parse Clx to find piece table
	fcClx := binary.LittleEndian.Uint32(wordDocData[0x1A2:0x1A6])
	lcbClx := binary.LittleEndian.Uint32(wordDocData[0x1A6:0x1AA])
	fmt.Printf("Clx: fc=%d lcb=%d\n", fcClx, lcbClx)

	clxData := tableData[fcClx : fcClx+lcbClx]

	// Find PlcPcd
	pos := uint32(0)
	for pos < uint32(len(clxData)) {
		if clxData[pos] == 0x01 {
			prcSize := binary.LittleEndian.Uint16(clxData[pos+1:])
			pos += 3 + uint32(prcSize)
			continue
		}
		if clxData[pos] == 0x02 {
			pos++
			break
		}
		pos++
	}

	plcPcdLen := binary.LittleEndian.Uint32(clxData[pos:])
	pos += 4
	plcPcdData := clxData[pos : pos+plcPcdLen]

	n := (plcPcdLen - 4) / 12
	fmt.Printf("Number of pieces: %d\n", n)

	// Read CPs
	for i := uint32(0); i <= n; i++ {
		cp := binary.LittleEndian.Uint32(plcPcdData[i*4:])
		if i == 0 || i == n {
			fmt.Printf("  CP[%d] = %d\n", i, cp)
		}
	}

	// Show last few pieces
	for i := n - 3; i < n; i++ {
		cpStart := binary.LittleEndian.Uint32(plcPcdData[i*4:])
		cpEnd := binary.LittleEndian.Uint32(plcPcdData[(i+1)*4:])
		pdStart := (n+1)*4 + i*8
		fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])
		isUnicode := fc&0x40000000 == 0
		fmt.Printf("  Piece[%d]: cp[%d-%d] fc=0x%08X unicode=%v\n", i, cpStart, cpEnd, fc, isUnicode)
	}
}
