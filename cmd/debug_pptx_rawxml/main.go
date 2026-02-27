package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for _, f := range zr.File {
		if f.Name == "ppt/slides/slide4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()

			// Find first <a:t> and show raw bytes
			content := string(data)
			// Find "背景" in the raw bytes
			// UTF-8 for 背 = e8 83 8c, 景 = e6 99 af
			target := []byte{0xe8, 0x83, 0x8c, 0xe6, 0x99, 0xaf}
			found := false
			for i := 0; i < len(data)-len(target); i++ {
				match := true
				for j := 0; j < len(target); j++ {
					if data[i+j] != target[j] {
						match = false
						break
					}
				}
				if match {
					found = true
					fmt.Printf("Found '背景' at byte offset %d\n", i)
					// Show surrounding bytes
					start := i - 20
					if start < 0 {
						start = 0
					}
					end := i + 40
					if end > len(data) {
						end = len(data)
					}
					fmt.Printf("  Context: %s\n", string(data[start:end]))
					break
				}
			}
			if !found {
				fmt.Println("'背景' NOT found in raw bytes!")
				// Show first <a:t> content
				for i := 0; i < len(content)-5; i++ {
					if content[i:i+5] == "<a:t>" {
						end := i + 5
						for end < len(content) && content[end] != '<' {
							end++
						}
						text := content[i+5 : end]
						fmt.Printf("First <a:t> content: %q\n", text)
						fmt.Printf("  hex: %x\n", []byte(text))
						// Show what this looks like decoded as different encodings
						break
					}
				}
			}

			// Also check the XML declaration
			fmt.Printf("\nFirst 100 bytes: %q\n", string(data[:100]))
			fmt.Printf("First 100 hex: %x\n", data[:100])
			break
		}
	}
}
