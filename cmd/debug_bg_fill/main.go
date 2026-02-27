package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
)

const recordHeaderSize = 8

func readRecordHeader(data []byte, offset uint32) (recVer uint16, recType uint16, recLen uint32, err error) {
	if int(offset)+8 > len(data) {
		return 0, 0, 0, fmt.Errorf("out of bounds")
	}
	verInst := binary.LittleEndian.Uint16(data[offset : offset+2])
	recType = binary.LittleEndian.Uint16(data[offset+2 : offset+4])
	recLen = binary.LittleEndian.Uint32(data[offset+4 : offset+8])
	recVer = verInst & 0x0F
	return recVer, recType, recLen, nil
}

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		panic(err)
	}
	defer f.Close()

	c, err := cfb.OpenReader(f)
	if err != nil {
		panic(err)
	}

	// Find PowerPoint Document stream
	var pptDoc []byte
	for _, entry := range c.GetEntries() {
		if entry.Name() == "PowerPoint Document" {
			pptDoc = entry.Data()
			break
		}
	}
	if pptDoc == nil {
		fmt.Println("PowerPoint Document stream not found")
		return
	}

	fmt.Printf("PPT Document size: %d bytes\n", len(pptDoc))

	// Find MainMasterContainer records and dump their background FOPT properties
	dataLen := uint32(len(pptDoc))
	offset := uint32(0)
	masterCount := 0

	for offset+8 <= dataLen {
		recVer, recType, recLen, err := readRecordHeader(pptDoc, offset)
		if err != nil {
			break
		}
		recDataStart := offset + 8
		recDataEnd := recDataStart + recLen
		if recDataEnd > dataLen {
			break
		}

		if recType == 0x03F8 { // rtMainMaster
			masterCount++
			fmt.Printf("\n=== MainMaster #%d at offset %d ===\n", masterCount, offset)
			// Scan for Drawing container
			scanForBgFOPT(pptDoc, recDataStart, recDataEnd, 0)
		}

		if recVer == 0xF {
			offset = recDataStart
		} else {
			offset = recDataEnd
		}
	}
}

func scanForBgFOPT(data []byte, start, end uint32, depth int) {
	pos := start
	for pos+8 <= end {
		recVer, recType, recLen, err := readRecordHeader(data, pos)
		if err != nil {
			break
		}
		recDataStart := pos + 8
		recDataEnd := recDataStart + recLen
		if recDataEnd > end {
			break
		}

		indent := ""
		for i := 0; i < depth; i++ {
			indent += "  "
		}

		// Print Drawing and SpContainer records
		if recType == 0xF000 { // rtDrawingGroup
			fmt.Printf("%sDrawingGroup at %d\n", indent, pos)
		}
		if recType == 0xF002 { // rtDrawing (DgContainer)
			fmt.Printf("%sDrawing at %d\n", indent, pos)
		}
		if recType == 0xF003 { // rtSpgrContainer
			fmt.Printf("%sSpgrContainer at %d\n", indent, pos)
		}
		if recType == 0xF004 { // rtSpContainer
			fmt.Printf("%sSpContainer at %d len=%d\n", indent, pos, recLen)
			// Check if this is a background shape
			scanSpContainer(data, recDataStart, recDataEnd, depth+1)
		}

		if recVer == 0xF {
			scanForBgFOPT(data, recDataStart, recDataEnd, depth+1)
			pos = recDataEnd
		} else {
			pos = recDataEnd
		}
	}
}

func scanSpContainer(data []byte, start, end uint32, depth int) {
	pos := start
	indent := ""
	for i := 0; i < depth; i++ {
		indent += "  "
	}

	for pos+8 <= end {
		_, recType, recLen, err := readRecordHeader(data, pos)
		if err != nil {
			break
		}
		recDataStart := pos + 8
		recDataEnd := recDataStart + recLen
		if recDataEnd > end {
			break
		}

		if recType == 0xF00A { // SpAtom
			if recLen >= 8 {
				spid := binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4])
				grfPersist := binary.LittleEndian.Uint32(data[recDataStart+4 : recDataStart+8])
				isBg := grfPersist&0x400 != 0
				fmt.Printf("%sSpAtom: spid=%d grfPersist=0x%08X isBg=%v\n", indent, spid, grfPersist, isBg)
			}
		}

		if recType == 0xF00B || recType == 0xF121 || recType == 0xF122 { // FOPT variants
			numProps := binary.LittleEndian.Uint16(data[pos:pos+2]) >> 4
			fmt.Printf("%sFOPT (type=0x%04X) numProps=%d:\n", indent, recType, numProps)
			for i := uint16(0); i < numProps && int(i)*6+6 <= int(recLen); i++ {
				propOff := recDataStart + uint32(i)*6
				propID := binary.LittleEndian.Uint16(data[propOff : propOff+2])
				propVal := binary.LittleEndian.Uint32(data[propOff+2 : propOff+6])
				basePropID := propID & 0x3FFF
				isComplex := propID&0x8000 != 0

				propName := fmt.Sprintf("0x%04X", basePropID)
				switch basePropID {
				case 0x0180:
					propName = "fillType"
				case 0x0181:
					propName = "fillColor"
				case 0x0182:
					propName = "fillOpacity"
				case 0x0183:
					propName = "fillBackColor"
				case 0x0184:
					propName = "fillBackOpacity"
				case 0x0185:
					propName = "fillCrMod"
				case 0x0186:
					propName = "fillBlip"
				case 0x0104:
					propName = "pib"
				case 0x01BF:
					propName = "fillBools"
				case 0x01C0:
					propName = "lineColor"
				case 0x01FF:
					propName = "lineBools"
				case 0x0187:
					propName = "fillBlipName"
				case 0x0188:
					propName = "fillBlipFlags"
				case 0x0189:
					propName = "fillWidth"
				case 0x018A:
					propName = "fillHeight"
				case 0x018B:
					propName = "fillAngle"
				case 0x018C:
					propName = "fillFocus"
				case 0x018D:
					propName = "fillToLeft"
				case 0x018E:
					propName = "fillToTop"
				case 0x018F:
					propName = "fillToRight"
				case 0x0190:
					propName = "fillToBottom"
				case 0x0191:
					propName = "fillRectLeft"
				case 0x0192:
					propName = "fillRectTop"
				case 0x0193:
					propName = "fillRectRight"
				case 0x0194:
					propName = "fillRectBottom"
				case 0x0195:
					propName = "fillDztype"
				case 0x0196:
					propName = "fillShadePreset"
				case 0x0197:
					propName = "fillShadeColors"
				case 0x0198:
					propName = "fillOriginX"
				case 0x0199:
					propName = "fillOriginY"
				case 0x019A:
					propName = "fillShapeOriginX"
				case 0x019B:
					propName = "fillShapeOriginY"
				case 0x019E:
					propName = "fillColorExt"
				case 0x019F:
					propName = "fillColorExtMod"
				case 0x01A0:
					propName = "fillBackColorExt"
				case 0x01A1:
					propName = "fillBackColorExtMod"
				}

				complexStr := ""
				if isComplex {
					complexStr = " [COMPLEX]"
				}
				fmt.Printf("%s  prop %s = 0x%08X (%d)%s\n", indent, propName, propVal, propVal, complexStr)
			}
		}

		pos = recDataEnd
	}
}
