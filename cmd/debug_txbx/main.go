package main

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"

	"github.com/shakinm/xlsReader/cfb"
	"github.com/shakinm/xlsReader/helpers"
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
	ccpEdn := binary.LittleEndian.Uint32(wordDocData[0x60:0x64])
	ccpTxbx := binary.LittleEndian.Uint32(wordDocData[0x64:0x68])
	ccpHdrTxbx := binary.LittleEndian.Uint32(wordDocData[0x68:0x6C])

	fmt.Printf("ccpText=%d ccpFtn=%d ccpHdd=%d ccpAtn=%d ccpEdn=%d ccpTxbx=%d ccpHdrTxbx=%d\n",
		ccpText, ccpFtn, ccpHdd, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx)

	// Text box text starts after main + footnote + header + annotation + endnote
	txbxStart := ccpText + ccpFtn + ccpHdd + ccpAtn + ccpEdn
	txbxEnd := txbxStart + ccpTxbx
	fmt.Printf("Text box area: CP %d to %d (%d chars)\n", txbxStart, txbxEnd, ccpTxbx)

	// Extract full text from piece table
	fcClx := binary.LittleEndian.Uint32(wordDocData[0x1A2:0x1A6])
	lcbClx := binary.LittleEndian.Uint32(wordDocData[0x1A6:0x1AA])
	clxData := tableData[fcClx : fcClx+lcbClx]

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

	fullText := ""
	for i := uint32(0); i < n; i++ {
		cpStart := binary.LittleEndian.Uint32(plcPcdData[i*4:])
		cpEnd := binary.LittleEndian.Uint32(plcPcdData[(i+1)*4:])
		pdStart := (n+1)*4 + i*8
		fc := binary.LittleEndian.Uint32(plcPcdData[pdStart+2:])
		isUnicode := fc&0x40000000 == 0

		charCount := cpEnd - cpStart
		var actualFC uint32
		if isUnicode {
			actualFC = fc
		} else {
			actualFC = (fc & ^uint32(0x40000000)) >> 1
		}

		var byteCount uint32
		if isUnicode {
			byteCount = charCount * 2
		} else {
			byteCount = charCount
		}

		end := actualFC + byteCount
		if uint64(end) > uint64(len(wordDocData)) {
			continue
		}

		fragment := wordDocData[actualFC:end]
		if isUnicode {
			fullText += helpers.DecodeUTF16LE(fragment)
		} else {
			fullText += helpers.DecodeANSI(fragment)
		}
		_ = cpStart
	}

	runes := []rune(fullText)
	fmt.Printf("Full text runes: %d\n", len(runes))

	// Show text box area
	if int(txbxStart) < len(runes) {
		end := int(txbxEnd)
		if end > len(runes) {
			end = len(runes)
		}
		fmt.Printf("\nText box text (CP %d-%d):\n", txbxStart, end)
		for i := int(txbxStart); i < end; i++ {
			r := runes[i]
			if r < 0x20 {
				fmt.Printf("[%02X]", r)
			} else {
				fmt.Printf("%c", r)
			}
		}
		fmt.Println()
	}

	// Check PlcfTxbxTxt - maps text box CPs to shapes
	// FibRgFcLcb97: fcPlcfTxbxTxt at offset 0x282, lcbPlcfTxbxTxt at 0x286
	// Actually these are in FibRgFcLcb97 which starts at different offsets
	// Let me check the FIB for text box placement info
	
	// PlcfSpa for text boxes in main document: fcPlcSpaMom/lcbPlcSpaMom
	fcPlcSpaMom := binary.LittleEndian.Uint32(wordDocData[0x1DA:0x1DE])
	lcbPlcSpaMom := binary.LittleEndian.Uint32(wordDocData[0x1DE:0x1E2])
	fmt.Printf("\nPlcSpaMom: fc=%d lcb=%d\n", fcPlcSpaMom, lcbPlcSpaMom)

	if lcbPlcSpaMom > 0 {
		spaData := tableData[fcPlcSpaMom : fcPlcSpaMom+lcbPlcSpaMom]
		// PlcfSpa: (n+1) CPs + n SPAs (26 bytes each)
		nSpa := (lcbPlcSpaMom - 4) / 30 // (n+1)*4 + n*26 = 4 + 30n
		fmt.Printf("Number of SPAs: %d\n", nSpa)
		for i := uint32(0); i < nSpa; i++ {
			cp := binary.LittleEndian.Uint32(spaData[i*4:])
			spaOff := (nSpa+1)*4 + i*26
			spid := binary.LittleEndian.Uint32(spaData[spaOff:])
			fmt.Printf("  SPA[%d]: cp=%d spid=%d\n", i, cp, spid)
		}
	}

	// Check PlcfTxbxTxt
	// In FibRgFcLcb97, fcPlcfTxbxTxt is at offset 0x282 from start of FIB
	// But the FIB layout is complex. Let me search for it.
	// FibRgFcLcb97 starts at offset 0x9A in the FIB
	// fcPlcfTxbxTxt is at FibRgFcLcb97 offset 0x1E8 (from MS-DOC 2.5.11)
	// So absolute offset = 0x9A + 0x1E8 = 0x282
	fcPlcfTxbxTxt := binary.LittleEndian.Uint32(wordDocData[0x282:0x286])
	lcbPlcfTxbxTxt := binary.LittleEndian.Uint32(wordDocData[0x286:0x28A])
	fmt.Printf("\nPlcfTxbxTxt: fc=%d lcb=%d\n", fcPlcfTxbxTxt, lcbPlcfTxbxTxt)

	if lcbPlcfTxbxTxt > 0 {
		txbxTxtData := tableData[fcPlcfTxbxTxt : fcPlcfTxbxTxt+lcbPlcfTxbxTxt]
		// PlcfTxbxTxt: (n+1) CPs + n TxbxTxt (4 bytes each? or variable)
		// Actually it's (n+1)*4 CPs + n*0 data = just CPs
		// Wait, PlcfTxbxTxt has (n+1) CPs and n FTXBXS (22 bytes each)
		// So: (n+1)*4 + n*22 = 4 + 26n
		nTxbx := (lcbPlcfTxbxTxt - 4) / 26
		fmt.Printf("Number of text boxes: %d\n", nTxbx)
		for i := uint32(0); i < nTxbx; i++ {
			cp := binary.LittleEndian.Uint32(txbxTxtData[i*4:])
			fmt.Printf("  TxbxTxt[%d]: cp=%d (absolute: %d)\n", i, cp, txbxStart+cp)
		}
		// Last CP
		lastCP := binary.LittleEndian.Uint32(txbxTxtData[nTxbx*4:])
		fmt.Printf("  TxbxTxt[last]: cp=%d (absolute: %d)\n", lastCP, txbxStart+lastCP)
	}

	// Check PlcfTxbxBkd - links text boxes to shapes
	// fcPlcTxbxBkd at FibRgFcLcb97 offset 0x1F0 => absolute 0x9A + 0x1F0 = 0x28A
	fcPlcTxbxBkd := binary.LittleEndian.Uint32(wordDocData[0x28A:0x28E])
	lcbPlcTxbxBkd := binary.LittleEndian.Uint32(wordDocData[0x28E:0x292])
	fmt.Printf("\nPlcTxbxBkd: fc=%d lcb=%d\n", fcPlcTxbxBkd, lcbPlcTxbxBkd)
	
	_ = tableData
}
