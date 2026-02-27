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
				fmt.Fprintf(os.Stderr, "Error opening stream: %v\n", err)
				os.Exit(1)
			}
			data := make([]byte, 100*1024*1024)
			n, _ := reader.Read(data)
			data = data[:n]
			fmt.Printf("PowerPoint Document: %d bytes\n", n)

			// Search for FOPT records that contain pib=14 (image index 13+1)
			// pib property ID = 0x0104
			// We look for the property in FOPT records
			for i := 0; i < len(data)-8; i++ {
				recVer := binary.LittleEndian.Uint16(data[i : i+2])
				recType := binary.LittleEndian.Uint16(data[i+2 : i+4])
				recLen := binary.LittleEndian.Uint32(data[i+4 : i+8])

				if recType != 0xF00B { // FOPT
					continue
				}
				numProps := recVer >> 4
				if numProps == 0 || numProps > 50 || recLen > 10000 {
					continue
				}

				// Check if this FOPT has pib=14
				propStart := i + 8
				hasPib14 := false
				for p := uint16(0); p < numProps && propStart+6 <= len(data); p++ {
					propID := binary.LittleEndian.Uint16(data[propStart : propStart+2])
					propVal := binary.LittleEndian.Uint32(data[propStart+2 : propStart+6])
					basePropID := propID & 0x3FFF
					if basePropID == 0x0104 && propVal == 14 { // pib=14 means imageIdx=13
						hasPib14 = true
						break
					}
					propStart += 6
				}

				if !hasPib14 {
					continue
				}

				fmt.Printf("\nFOPT at offset %d with pib=14 (watermark image):\n", i)
				fmt.Printf("  numProps=%d recLen=%d\n", numProps, recLen)
				propStart = i + 8
				for p := uint16(0); p < numProps && propStart+6 <= len(data); p++ {
					propID := binary.LittleEndian.Uint16(data[propStart : propStart+2])
					propVal := binary.LittleEndian.Uint32(data[propStart+2 : propStart+6])
					basePropID := propID & 0x3FFF
					isComplex := propID&0x8000 != 0
					propName := propIDName(basePropID)
					fmt.Printf("  prop[%d]: id=0x%04X (%s) val=0x%08X (%d) complex=%v\n",
						p, basePropID, propName, propVal, propVal, isComplex)
					propStart += 6
				}
			}
		}
	}
}

func propIDName(id uint16) string {
	names := map[uint16]string{
		0x0004: "rotation",
		0x007F: "protectionBools",
		0x0080: "lTxid",
		0x0081: "dxTextLeft",
		0x0082: "dyTextTop",
		0x0083: "dxTextRight",
		0x0084: "dyTextBottom",
		0x0085: "wrapText",
		0x0086: "unused86",
		0x0087: "anchorText",
		0x00BF: "textBools",
		0x0100: "cropFromTop",
		0x0101: "cropFromBottom",
		0x0102: "cropFromLeft",
		0x0103: "cropFromRight",
		0x0104: "pib",
		0x0105: "pibName",
		0x0106: "pibFlags",
		0x0107: "pictureTransparent",
		0x0108: "pictureContrast",
		0x0109: "pictureBrightness",
		0x010A: "pictureGamma",
		0x010B: "pictureId",
		0x010C: "pictureDblCrMod",
		0x010D: "pictureFillCrMod",
		0x010E: "pictureLineCrMod",
		0x010F: "pibPrint",
		0x0110: "pibPrintName",
		0x0111: "pibPrintFlags",
		0x013F: "pictureBools",
		0x0140: "geoLeft",
		0x0141: "geoTop",
		0x0142: "geoRight",
		0x0143: "geoBottom",
		0x0144: "shapePath",
		0x0145: "pVertices",
		0x0146: "pSegmentInfo",
		0x0180: "fillType",
		0x0181: "fillColor",
		0x0182: "fillOpacity",
		0x0183: "fillBackColor",
		0x0184: "fillBackOpacity",
		0x0185: "fillCrMod",
		0x0186: "fillBlip",
		0x0187: "fillBlipName",
		0x0188: "fillBlipFlags",
		0x0189: "fillWidth",
		0x018A: "fillHeight",
		0x018B: "fillAngle",
		0x018C: "fillFocus",
		0x018D: "fillToLeft",
		0x018E: "fillToTop",
		0x018F: "fillToRight",
		0x0190: "fillToBottom",
		0x0191: "fillRectLeft",
		0x0192: "fillRectTop",
		0x0193: "fillRectRight",
		0x0194: "fillRectBottom",
		0x0195: "fillDztype",
		0x0196: "fillShadePreset",
		0x0197: "fillShadeColors",
		0x0198: "fillOriginX",
		0x0199: "fillOriginY",
		0x019A: "fillShapeOriginX",
		0x019B: "fillShapeOriginY",
		0x019C: "fillShadeType",
		0x01BF: "fillBools",
		0x01C0: "lineColor",
		0x01C1: "lineOpacity",
		0x01C2: "lineBackColor",
		0x01C3: "lineCrMod",
		0x01C4: "lineType",
		0x01C5: "lineFillBlip",
		0x01C6: "lineFillBlipName",
		0x01C7: "lineFillBlipFlags",
		0x01C8: "lineFillWidth",
		0x01C9: "lineFillHeight",
		0x01CA: "lineFillDztype",
		0x01CB: "lineWidth",
		0x01CC: "lineMiterLimit",
		0x01CD: "lineStyle",
		0x01CE: "lineDashing",
		0x01CF: "lineDashStyle",
		0x01D0: "lineStartArrowhead",
		0x01D1: "lineEndArrowhead",
		0x01D2: "lineStartArrowWidth",
		0x01D3: "lineStartArrowLength",
		0x01D4: "lineEndArrowWidth",
		0x01D5: "lineEndArrowLength",
		0x01D6: "lineJoinStyle",
		0x01D7: "lineEndCapStyle",
		0x01FF: "lineBools",
		0x0301: "shadowColor",
		0x0302: "shadowHighlight",
		0x0303: "shadowCrMod",
		0x0304: "shadowOpacity",
		0x0305: "shadowOffsetX",
		0x0306: "shadowOffsetY",
		0x033F: "shadowBools",
		0x0380: "perspectiveType",
		0x03BF: "perspectiveBools",
		0x0500: "cxk",
		0x053F: "groupBools",
	}
	if name, ok := names[id]; ok {
		return name
	}
	return fmt.Sprintf("unknown_0x%04X", id)
}
