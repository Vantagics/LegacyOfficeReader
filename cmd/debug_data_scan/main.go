package main

import (
	"encoding/binary"
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/cfb"
)

func main() {
	adaptor, err := cfb.OpenFile("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer adaptor.CloseFile()

	var root, dataDir *cfb.Directory
	for _, dir := range adaptor.GetDirs() {
		switch dir.Name() {
		case "Root Entry":
			root = dir
		case "Data":
			dataDir = dir
		}
	}

	dReader, _ := adaptor.OpenObject(dataDir, root)
	dSize := binary.LittleEndian.Uint32(dataDir.StreamSize[:])
	data := make([]byte, dSize)
	dReader.Read(data)

	fmt.Printf("Data stream size: %d bytes\n\n", len(data))

	// Scan for all OfficeArt records in the Data stream
	fmt.Println("=== All OfficeArt records ===")
	pos := 0
	for pos+8 <= len(data) {
		vi := binary.LittleEndian.Uint16(data[pos:])
		rt := binary.LittleEndian.Uint16(data[pos+2:])
		rl := binary.LittleEndian.Uint32(data[pos+4:])
		ver := vi & 0x0F
		inst := vi >> 4

		// Check if this looks like a valid OfficeArt record
		if rt >= 0xF000 && rt <= 0xF200 && rl > 0 && uint32(pos)+8+rl <= uint32(len(data)) {
			recName := fmt.Sprintf("0x%04X", rt)
			switch rt {
			case 0xF000:
				recName = "DggContainer"
			case 0xF001:
				recName = "BStoreContainer"
			case 0xF002:
				recName = "DgContainer"
			case 0xF003:
				recName = "SpgrContainer"
			case 0xF004:
				recName = "SpContainer"
			case 0xF006:
				recName = "Dgg"
			case 0xF007:
				recName = "BSE"
			case 0xF008:
				recName = "Dg"
			case 0xF009:
				recName = "Spgr"
			case 0xF00A:
				recName = "Fsp"
			case 0xF00B:
				recName = "Fopt"
			case 0xF01A:
				recName = "BlipEMF"
			case 0xF01B:
				recName = "BlipWMF"
			case 0xF01E:
				recName = "BlipPNG"
			case 0xF01D:
				recName = "BlipJPEG"
			}
			fmt.Printf("@%d: %s ver=%d inst=%d len=%d\n", pos, recName, ver, inst, rl)

			if ver == 0x0F {
				// Container - don't skip, let inner records be found
				pos += 8
				continue
			}
			pos += 8 + int(rl)
			continue
		}

		// Check for PNG signature
		if pos+8 <= len(data) && data[pos] == 0x89 && data[pos+1] == 'P' && data[pos+2] == 'N' && data[pos+3] == 'G' {
			fmt.Printf("@%d: PNG signature found!\n", pos)
		}

		pos++
	}

	// Also check what's at the PICF offsets
	fmt.Println("\n=== PICF structures ===")
	for _, offset := range []int{4085, 111211, 457622} {
		if offset+68 > len(data) {
			continue
		}
		// PICF structure: first 4 bytes are lcb (total size including header)
		lcb := binary.LittleEndian.Uint32(data[offset:])
		cbHeader := binary.LittleEndian.Uint16(data[offset+4:])
		// mm (mapping mode) at offset 6
		mm := binary.LittleEndian.Uint16(data[offset+6:])
		// xExt, yExt at offset 8, 10
		xExt := binary.LittleEndian.Uint16(data[offset+8:])
		yExt := binary.LittleEndian.Uint16(data[offset+10:])

		fmt.Printf("PICF@%d: lcb=%d, cbHeader=%d, mm=%d, xExt=%d, yExt=%d\n",
			offset, lcb, cbHeader, mm, xExt, yExt)

		// The total PICF structure size is lcb bytes
		// After cbHeader bytes comes the SpContainer
		// After the SpContainer, there might be image data
		spOff := offset + int(cbHeader)
		if spOff+8 <= len(data) {
			srl := binary.LittleEndian.Uint32(data[spOff+4:])
			_ = binary.LittleEndian.Uint16(data[spOff+2:]) // recType
			spEnd := spOff + 8 + int(srl)
			fmt.Printf("  SpContainer at %d, len=%d, ends at %d\n", spOff, srl, spEnd)

			// Check what's after the SpContainer but still within lcb
			picfEnd := offset + int(lcb)
			if spEnd < picfEnd && spEnd+8 <= len(data) {
				afterRt := binary.LittleEndian.Uint16(data[spEnd+2:])
				afterRl := binary.LittleEndian.Uint32(data[spEnd+4:])
				fmt.Printf("  After SpContainer at %d: recType=0x%04X, len=%d (remaining=%d)\n",
					spEnd, afterRt, afterRl, picfEnd-spEnd)

				// Check if it's a blip
				if afterRt >= 0xF018 && afterRt <= 0xF029 {
					fmt.Printf("  → EMBEDDED BLIP found! type=0x%04X, size=%d\n", afterRt, afterRl)
				}
			}
		}
	}
}
