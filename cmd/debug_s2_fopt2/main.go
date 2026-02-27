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
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
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

			// Search for FOPT records that have lineColor=0x08000001
			count := 0
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

				propStart := i + 8
				for p := uint16(0); p < numProps && propStart+6 <= len(data); p++ {
					propID := binary.LittleEndian.Uint16(data[propStart : propStart+2])
					propVal := binary.LittleEndian.Uint32(data[propStart+2 : propStart+6])
					basePropID := propID & 0x3FFF
					if basePropID == 0x01C0 && propVal == 0x08000001 {
						count++
						if count <= 3 {
							fmt.Printf("\nFOPT at offset %d with lineColor=0x08000001:\n", i)
							// Print all props
							ps := i + 8
							for pp := uint16(0); pp < numProps && ps+6 <= len(data); pp++ {
								pid := binary.LittleEndian.Uint16(data[ps : ps+2])
								pv := binary.LittleEndian.Uint32(data[ps+2 : ps+6])
								bpid := pid & 0x3FFF
								ic := pid&0x8000 != 0
								fmt.Printf("  prop[%d]: 0x%04X = 0x%08X complex=%v\n", pp, bpid, pv, ic)
								ps += 6
							}
						}
					}
					propStart += 6
				}
			}
			fmt.Printf("\nTotal FOPT records with lineColor=0x08000001: %d\n", count)
		}
	}
}
