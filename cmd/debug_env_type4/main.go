package main

import (
	"encoding/binary"
	"fmt"

	"github.com/shakinm/xlsReader/cfb"
)

const recordHeaderSize = 8

type rh struct {
	vi  uint16
	rt  uint16
	len uint32
}

func read(data []byte, off uint32) (rh, error) {
	if uint32(len(data)) < off+8 {
		return rh{}, fmt.Errorf("eof")
	}
	return rh{
		vi:  binary.LittleEndian.Uint16(data[off:]),
		rt:  binary.LittleEndian.Uint16(data[off+2:]),
		len: binary.LittleEndian.Uint32(data[off+4:]),
	}, nil
}

func (r rh) ver() uint16  { return r.vi & 0x0F }
func (r rh) inst() uint16 { return r.vi >> 4 }

func main() {
	adaptor, err := cfb.OpenFile("testfie/test.ppt")
	if err != nil {
		panic(err)
	}
	defer adaptor.CloseFile()

	var root, pptDoc *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "PowerPoint Document":
			pptDoc = dir
		}
	}

	reader, _ := adaptor.OpenObject(pptDoc, root)
	size := binary.LittleEndian.Uint32(pptDoc.StreamSize[:])
	data := make([]byte, size)
	reader.Read(data)

	// Find Environment (0x03F2), then find TextMasterStyleAtom instance 4
	dataLen := uint32(len(data))
	off := uint32(0)
	for off+8 <= dataLen {
		r, err := read(data, off)
		if err != nil {
			break
		}
		ds := off + 8
		de := ds + r.len
		if de > dataLen {
			break
		}
		if r.rt == 0x03F2 {
			fmt.Println("Found Environment, scanning recursively...")
			scanEnv(data, ds, de, 0)
			return
		}
		if r.ver() == 0xF {
			off = ds
		} else {
			off = de
		}
	}
	fmt.Println("Environment not found")
}

func scanEnv(data []byte, start, end uint32, depth int) {
	off := start
	for off+8 <= end {
		r, err := read(data, off)
		if err != nil {
			break
		}
		ds := off + 8
		de := ds + r.len
		if de > end {
			break
		}

		prefix := ""
		for i := 0; i < depth; i++ {
			prefix += "  "
		}

		if r.rt == 0x0FA3 {
			fmt.Printf("%sTextMasterStyleAtom inst=%d len=%d\n", prefix, r.inst(), r.len)
			// Dump raw hex
			fmt.Printf("%s  hex: ", prefix)
			for i := ds; i < de && i < ds+120; i++ {
				fmt.Printf("%02X ", data[i])
			}
			fmt.Println()

			// Try to parse
			if r.len >= 2 {
				numLevels := binary.LittleEndian.Uint16(data[ds : ds+2])
				fmt.Printf("%s  numLevels=%d\n", prefix, numLevels)
				pos := ds + 2
				for level := 0; level < int(numLevels) && pos < de; level++ {
					if pos+2 > de {
						break
					}
					indentLevel := binary.LittleEndian.Uint16(data[pos : pos+2])
					pos += 2
					fmt.Printf("%s  Level %d (indent=%d):\n", prefix, level, indentLevel)

					// Parse paragraph mask
					if pos+4 > de {
						break
					}
					paraMask := binary.LittleEndian.Uint32(data[pos : pos+4])
					fmt.Printf("%s    paraMask=0x%08X\n", prefix, paraMask)
					pos += 4
					// Skip paragraph props based on mask
					pos = skipParaProps(data, pos, de, paraMask)

					// Parse character mask
					if pos+4 > de {
						break
					}
					charMask := binary.LittleEndian.Uint32(data[pos : pos+4])
					fmt.Printf("%s    charMask=0x%08X\n", prefix, charMask)
					pos += 4
					// Parse character props
					parseCharProps(data, pos, de, charMask, prefix+"    ")
					pos = skipCharProps(data, pos, de, charMask)
				}
			}
		}

		if r.ver() == 0xF {
			scanEnv(data, ds, de, depth+1)
			off = de
		} else {
			off = de
		}
	}
}

func skipParaProps(data []byte, pos, end uint32, mask uint32) uint32 {
	if mask&0x000F != 0 { // bullet flags
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x0010 != 0 { // bulletChar
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x0020 != 0 { // bulletFont
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x0040 != 0 { // bulletSize
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x0080 != 0 { // bulletColor
		if pos+4 <= end {
			pos += 4
		}
	}
	if mask&0x0800 != 0 { // align
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x1000 != 0 { // lineSpacing
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x2000 != 0 { // spaceBefore
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x4000 != 0 { // spaceAfter
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x8000 != 0 { // leftMargin
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x10000 != 0 { // indent
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x100000 != 0 { // defaultTab
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x200000 != 0 { // tabStops
		if pos+2 <= end {
			count := binary.LittleEndian.Uint16(data[pos : pos+2])
			pos += 2 + uint32(count)*4
		}
	}
	if mask&0x400000 != 0 { // fontAlign
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x3800000 != 0 { // charWrap/wordWrap/overflow
		if pos+2 <= end {
			pos += 2
		}
	}
	if mask&0x4000000 != 0 { // textDirection
		if pos+2 <= end {
			pos += 2
		}
	}
	return pos
}

func parseCharProps(data []byte, pos, end uint32, mask uint32, prefix string) {
	if mask&0x0001FFFF != 0 { // style bits
		if pos+2 <= end {
			flags := binary.LittleEndian.Uint16(data[pos : pos+2])
			fmt.Printf("%sflags=0x%04X (bold=%v italic=%v underline=%v)\n", prefix, flags, flags&1 != 0, flags&2 != 0, flags&4 != 0)
			pos += 2
		}
	}
	if mask&0x00020000 != 0 { // typeface
		if pos+2 <= end {
			fontIdx := binary.LittleEndian.Uint16(data[pos : pos+2])
			fmt.Printf("%sfontIdx=%d\n", prefix, fontIdx)
			pos += 2
		}
	}
	if mask&0x00040000 != 0 { // oldEATypeface
		pos += 2
	}
	if mask&0x00080000 != 0 { // ansiTypeface
		pos += 2
	}
	if mask&0x00100000 != 0 { // symbolTypeface
		pos += 2
	}
	if mask&0x00200000 != 0 { // size
		if pos+2 <= end {
			sz := binary.LittleEndian.Uint16(data[pos : pos+2])
			fmt.Printf("%sfontSize=%d (centipoints=%d)\n", prefix, sz, sz*100)
			pos += 2
		}
	}
	if mask&0x00400000 != 0 { // color
		if pos+4 <= end {
			colorVal := binary.LittleEndian.Uint32(data[pos : pos+4])
			r := uint8(colorVal & 0xFF)
			g := uint8((colorVal >> 8) & 0xFF)
			b := uint8((colorVal >> 16) & 0xFF)
			fmt.Printf("%scolor=%02X%02X%02X colorRaw=0x%08X\n", prefix, r, g, b, colorVal)
			pos += 4
		}
	}
	_ = pos
}

func skipCharProps(data []byte, pos, end uint32, mask uint32) uint32 {
	if mask&0x0001FFFF != 0 {
		pos += 2
	}
	if mask&0x00020000 != 0 {
		pos += 2
	}
	if mask&0x00040000 != 0 {
		pos += 2
	}
	if mask&0x00080000 != 0 {
		pos += 2
	}
	if mask&0x00100000 != 0 {
		pos += 2
	}
	if mask&0x00200000 != 0 {
		pos += 2
	}
	if mask&0x00400000 != 0 {
		pos += 4
	}
	if mask&0x00800000 != 0 { // position
		pos += 2
	}
	return pos
}
