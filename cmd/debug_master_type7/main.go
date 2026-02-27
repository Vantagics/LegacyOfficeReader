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

const rtMainMaster = 0x03F8

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

		if r.rt == rtMainMaster {
			fmt.Printf("Found MainMaster at offset %d\n", off)
			scanForTMS(data, ds, de, 0)
		}

		if r.ver() == 0xF {
			off = ds
		} else {
			off = de
		}
	}
}

func scanForTMS(data []byte, start, end uint32, depth int) {
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

		if r.rt == 0x0FA3 {
			prefix := ""
			for i := 0; i < depth; i++ {
				prefix += "  "
			}
			fmt.Printf("%sTextMasterStyleAtom inst=%d len=%d\n", prefix, r.inst(), r.len)
			fmt.Printf("%s  hex: ", prefix)
			maxBytes := de
			if maxBytes > ds+200 {
				maxBytes = ds + 200
			}
			for i := ds; i < maxBytes; i++ {
				fmt.Printf("%02X ", data[i])
			}
			fmt.Println()
		}

		if r.ver() == 0xF {
			scanForTMS(data, ds, de, depth+1)
			off = de
		} else {
			off = de
		}
	}
}
