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
	var root, dataDir *cfb.Directory
	for _, d := range dirs {
		name := d.Name()
		if name == "Root Entry" { root = d }
		if name == "Data" { dataDir = d }
	}

	if dataDir == nil {
		fmt.Println("No Data stream")
		return
	}

	dReader, _ := c.OpenObject(dataDir, root)
	data, _ := io.ReadAll(dReader)
	fmt.Printf("Data stream: %d bytes\n", len(data))

	// Scan for all SpContainers and extract their properties
	dataLen := uint32(len(data))
	for i := uint32(0); i+8 <= dataLen; i++ {
		verInst := binary.LittleEndian.Uint16(data[i : i+2])
		recVer := verInst & 0x0F
		recType := binary.LittleEndian.Uint16(data[i+2 : i+4])
		recLen := binary.LittleEndian.Uint32(data[i+4 : i+8])

		if recType == 0xF004 && recVer == 0xF && recLen > 0 && i+8+recLen <= dataLen {
			// SpContainer found
			spid, pib, txid, hasClientTextbox := parseSpContainer(data, i+8, i+8+recLen)
			if spid != 0 {
				fmt.Printf("SpContainer at %d: spid=%d pib=%d txid=%d hasClientTextbox=%v\n",
					i, spid, pib, txid, hasClientTextbox)
			}
			i += 7 + recLen
		}
	}
}

func parseSpContainer(data []byte, offset, limit uint32) (spid, pib, txid uint32, hasClientTextbox bool) {
	for offset+8 <= limit {
		verInst := binary.LittleEndian.Uint16(data[offset : offset+2])
		recVer := verInst & 0x0F
		recInst := verInst >> 4
		recType := binary.LittleEndian.Uint16(data[offset+2 : offset+4])
		recLen := binary.LittleEndian.Uint32(data[offset+4 : offset+8])

		childEnd := offset + 8 + recLen
		if childEnd > limit {
			break
		}

		if recType == 0xF00A && recLen >= 8 { // Sp
			spid = binary.LittleEndian.Uint32(data[offset+8:])
		}

		if recType == 0xF00D { // ClientTextbox
			hasClientTextbox = true
		}

		if recType == 0xF00B && recLen > 0 { // OPT
			nProps := recInst
			propOff := offset + 8
			for p := uint16(0); p < nProps && propOff+6 <= childEnd; p++ {
				propID := binary.LittleEndian.Uint16(data[propOff:])
				propVal := binary.LittleEndian.Uint32(data[propOff+2:])
				pid := propID & 0x3FFF
				if pid == 0x0104 { // pib
					pib = propVal
				}
				if pid == 0x0080 { // lTxid
					txid = propVal
				}
				propOff += 6
			}
		}

		if recVer == 0xF {
			s, p, t, ct := parseSpContainer(data, offset+8, childEnd)
			if s != 0 { spid = s }
			if p != 0 { pib = p }
			if t != 0 { txid = t }
			if ct { hasClientTextbox = true }
		}

		offset = childEnd
	}
	return
}
