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

	// For each PicLocation, dump the full SpContainer contents
	for _, offset := range []int{4085, 111211, 457622} {
		fmt.Printf("\n=== PicLocation %d ===\n", offset)
		if offset+6 > len(data) {
			continue
		}
		cbHeader := int(binary.LittleEndian.Uint16(data[offset+4:]))
		fmt.Printf("PICF cbHeader: %d\n", cbHeader)

		spOff := offset + cbHeader
		if spOff+8 > len(data) {
			continue
		}

		vi := binary.LittleEndian.Uint16(data[spOff:])
		rt := binary.LittleEndian.Uint16(data[spOff+2:])
		rl := binary.LittleEndian.Uint32(data[spOff+4:])
		fmt.Printf("SpContainer: verInst=0x%04X, recType=0x%04X, recLen=%d\n", vi, rt, rl)

		// Dump all records inside the SpContainer
		end := spOff + 8 + int(rl)
		if end > len(data) {
			end = len(data)
		}
		pos := spOff + 8
		for pos+8 <= end {
			cvi := binary.LittleEndian.Uint16(data[pos:])
			crt := binary.LittleEndian.Uint16(data[pos+2:])
			crl := binary.LittleEndian.Uint32(data[pos+4:])
			cinst := cvi >> 4
			cver := cvi & 0x0F

			childEnd := pos + 8 + int(crl)
			if childEnd > end {
				childEnd = end
			}

			recName := fmt.Sprintf("0x%04X", crt)
			switch crt {
			case 0xF00A:
				recName = "Fsp"
			case 0xF00B:
				recName = "Fopt"
			case 0xF00D:
				recName = "ClientAnchor"
			case 0xF010:
				recName = "ClientData"
			case 0xF01A:
				recName = "BlipEMF"
			case 0xF01B:
				recName = "BlipWMF"
			case 0xF01C:
				recName = "BlipPICT"
			case 0xF01D:
				recName = "BlipJPEG"
			case 0xF01E:
				recName = "BlipPNG"
			case 0xF01F:
				recName = "BlipDIB"
			case 0xF029:
				recName = "BlipTIFF"
			case 0xF007:
				recName = "BSE"
			case 0xF11E:
				recName = "SplitMenuColors"
			case 0xF122:
				recName = "TertiaryFopt"
			}

			fmt.Printf("  %s: ver=%d, inst=%d, len=%d\n", recName, cver, cinst, crl)

			if crt == 0xF00A && childEnd-pos >= 16 {
				spid := binary.LittleEndian.Uint32(data[pos+8:])
				flags := binary.LittleEndian.Uint32(data[pos+12:])
				fmt.Printf("    SPID=%d, flags=0x%08X\n", spid, flags)
			}

			if crt == 0xF00B { // Fopt
				recData := data[pos+8 : childEnd]
				for p := uint16(0); p < cinst; p++ {
					off := int(p) * 6
					if off+6 > len(recData) {
						break
					}
					propID := binary.LittleEndian.Uint16(recData[off:])
					propVal := binary.LittleEndian.Uint32(recData[off+2:])
					pid := propID & 0x3FFF
					isComplex := (propID & 0x8000) != 0
					isBid := (propID & 0x4000) != 0
					pidName := fmt.Sprintf("0x%04X", pid)
					switch pid {
					case 0x0004:
						pidName = "rotation"
					case 0x007F:
						pidName = "lockAgainstGrouping"
					case 0x0080:
						pidName = "lTxid"
					case 0x0081:
						pidName = "dxTextLeft"
					case 0x0100:
						pidName = "cropFromTop"
					case 0x0104:
						pidName = "pib"
					case 0x0105:
						pidName = "pibName"
					case 0x0140:
						pidName = "fillType"
					case 0x0181:
						pidName = "fNoFillHitTest"
					case 0x01C0:
						pidName = "lineColor"
					case 0x01CB:
						pidName = "lineWidth"
					case 0x01FF:
						pidName = "fNoLineDrawDash"
					}
					fmt.Printf("    prop %s = %d (0x%08X) complex=%v bid=%v\n", pidName, propVal, propVal, isComplex, isBid)
				}
			}

			// Check for embedded blip
			if crt >= 0xF018 && crt <= 0xF029 {
				fmt.Printf("    EMBEDDED BLIP! size=%d bytes\n", crl)
			}

			pos = childEnd
		}
	}
}
