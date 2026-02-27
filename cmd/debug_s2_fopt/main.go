package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	c, err := cfb.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "CFB error: %v\n", err)
		os.Exit(1)
	}

	dirs := c.GetDirs()
	var root *cfb.Directory
	for _, d := range dirs {
		if d.Name() == "Root Entry" {
			root = d
			break
		}
	}

	for _, d := range dirs {
		if d.Name() == "PowerPoint Document" {
			reader, err := c.OpenObject(d, root)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}
			data := make([]byte, 100*1024*1024)
			n, _ := reader.Read(data)
			data = data[:n]

			// Search for FOPT records with lineColor=0C0D0E (0x0E0D0C in BGR)
			// and lineWidth=12700 (0x319C)
			targetLineColor := uint32(0x0E0D0C) // BGR for 0C0D0E
			targetLineWidth := uint32(12700)

			for i := 0; i < len(data)-8; i++ {
				recVer := binary.LittleEndian.Uint16(data[i : i+2])
				recType := binary.LittleEndian.Uint16(data[i+2 : i+4])
				recLen := binary.LittleEndian.Uint32(data[i+4 : i+8])

				if recType != 0xF00B {
					continue
				}
				numProps := recVer >> 4
				if numProps == 0 || numProps > 50 || recLen > 10000 {
					continue
				}

				// Check if this FOPT has lineColor=0C0D0E and lineWidth=12700
				propStart := i + 8
				hasTargetLine := false
				_ = hasTargetLine
				hasTargetWidth := false
				for p := uint16(0); p < numProps && propStart+6 <= len(data); p++ {
					propID := binary.LittleEndian.Uint16(data[propStart : propStart+2])
					propVal := binary.LittleEndian.Uint32(data[propStart+2 : propStart+6])
					basePropID := propID & 0x3FFF
					if basePropID == 0x01C0 && propVal == targetLineColor {
						hasTargetLine = true
					}
					if basePropID == 0x01CB && propVal == targetLineWidth {
						hasTargetWidth = true
					}
					propStart += 6
				}

				if !hasTargetWidth {
					continue
				}

				fmt.Printf("\nFOPT at offset %d with line=0C0D0E width=12700:\n", i)
				propStart = i + 8
				for p := uint16(0); p < numProps && propStart+6 <= len(data); p++ {
					propID := binary.LittleEndian.Uint16(data[propStart : propStart+2])
					propVal := binary.LittleEndian.Uint32(data[propStart+2 : propStart+6])
					basePropID := propID & 0x3FFF
					isComplex := propID&0x8000 != 0
					name := propName(basePropID)
					fmt.Printf("  prop[%d]: 0x%04X (%s) = 0x%08X (%d) complex=%v\n",
						p, basePropID, name, propVal, propVal, isComplex)
					propStart += 6
				}
				// Only show first match
				return
			}
		}
	}
}

func propName(id uint16) string {
	names := map[uint16]string{
		0x0004: "rotation",
		0x007F: "protectionBools",
		0x0080: "lTxid",
		0x0081: "dxTextLeft",
		0x0082: "dyTextTop",
		0x0083: "dxTextRight",
		0x0084: "dyTextBottom",
		0x0085: "wrapText",
		0x0086: "anchorText",
		0x00BF: "textBools",
		0x0104: "pib",
		0x013F: "pictureBools",
		0x0180: "fillType",
		0x0181: "fillColor",
		0x0182: "fillOpacity",
		0x01BF: "fillBools",
		0x01C0: "lineColor",
		0x01C1: "lineOpacity",
		0x01CB: "lineWidth",
		0x01CD: "lineStyle",
		0x01CE: "lineDashing",
		0x01D0: "lineStartArrowhead",
		0x01D1: "lineEndArrowhead",
		0x01FF: "lineBools",
		0x033F: "shadowBools",
		0x03BF: "perspectiveBools",
		0x053F: "groupBools",
	}
	if n, ok := names[id]; ok {
		return n
	}
	return fmt.Sprintf("unknown_0x%04X", id)
}
