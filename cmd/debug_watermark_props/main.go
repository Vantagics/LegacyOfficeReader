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

	c, err := cfb.OpenCfb(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "CFB error: %v\n", err)
		os.Exit(1)
	}

	// Read PowerPoint Document stream
	for _, entry := range c.GetEntries() {
		if entry.Name() == "PowerPoint Document" {
			data := entry.Data()
			fmt.Printf("PowerPoint Document: %d bytes\n", len(data))

			// Search for the watermark image shape by looking for its position
			// The watermark is at position (6026150, 4087812) = (0x5BF1C6, 0x3E6B24)
			// In EMU, these are stored as int32 LE
			targetX := int32(6026150)
			targetY := int32(4087812)

			xBytes := make([]byte, 4)
			yBytes := make([]byte, 4)
			binary.LittleEndian.PutUint32(xBytes, uint32(targetX))
			binary.LittleEndian.PutUint32(yBytes, uint32(targetY))

			// Search for the position bytes in the data
			for i := 0; i < len(data)-8; i++ {
				if data[i] == xBytes[0] && data[i+1] == xBytes[1] && data[i+2] == xBytes[2] && data[i+3] == xBytes[3] {
					// Check if Y follows
					if i+8 <= len(data) && data[i+4] == yBytes[0] && data[i+5] == yBytes[1] && data[i+6] == yBytes[2] && data[i+7] == yBytes[3] {
						fmt.Printf("Found position at offset %d (0x%X)\n", i, i)
						// Look backwards for FOPT record
						// Show surrounding bytes
						start := i - 200
						if start < 0 {
							start = 0
						}
						end := i + 200
						if end > len(data) {
							end = len(data)
						}

						// Look for FOPT (recType=0xF00B) near this position
						for j := start; j < i; j++ {
							if j+8 <= len(data) {
								recVer := binary.LittleEndian.Uint16(data[j : j+2])
								recType := binary.LittleEndian.Uint16(data[j+2 : j+4])
								recLen := binary.LittleEndian.Uint32(data[j+4 : j+8])
								if recType == 0xF00B && recLen > 0 && recLen < 10000 {
									numProps := recVer >> 4
									fmt.Printf("  FOPT at offset %d: numProps=%d len=%d\n", j, numProps, recLen)
									// Parse properties
									propStart := j + 8
									for p := uint16(0); p < numProps && propStart+6 <= len(data); p++ {
										propID := binary.LittleEndian.Uint16(data[propStart : propStart+2])
										propVal := binary.LittleEndian.Uint32(data[propStart+2 : propStart+6])
										basePropID := propID & 0x3FFF
										isComplex := propID&0x8000 != 0
										fmt.Printf("    prop[%d]: id=0x%04X (base=0x%04X) val=0x%08X complex=%v\n",
											p, propID, basePropID, propVal, isComplex)
										propStart += 6
									}
								}
							}
						}
					}
				}
			}
		}
	}
}
