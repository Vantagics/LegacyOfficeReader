package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
)

func main() {
	r1, err := zip.OpenReader("testfie/test_new7.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open new: %v\n", err)
		os.Exit(1)
	}
	defer r1.Close()

	r2, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "open ref: %v\n", err)
		os.Exit(1)
	}
	defer r2.Close()

	newFiles := make(map[string]string)
	for _, f := range r1.File {
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		newFiles[f.Name] = string(data)
	}

	refFiles := make(map[string]string)
	for _, f := range r2.File {
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()
		refFiles[f.Name] = string(data)
	}

	fmt.Println("=== Files in NEW but not in REF ===")
	for name := range newFiles {
		if _, ok := refFiles[name]; !ok {
			fmt.Printf("  + %s\n", name)
		}
	}

	fmt.Println("\n=== Files in REF but not in NEW ===")
	for name := range refFiles {
		if _, ok := newFiles[name]; !ok {
			fmt.Printf("  - %s\n", name)
		}
	}

	fmt.Println("\n=== Files that differ ===")
	for name, newData := range newFiles {
		if refData, ok := refFiles[name]; ok {
			if newData != refData {
				fmt.Printf("  ~ %s (new=%d bytes, ref=%d bytes)\n", name, len(newData), len(refData))
			}
		}
	}

	fmt.Println("\n=== Files that are identical ===")
	for name, newData := range newFiles {
		if refData, ok := refFiles[name]; ok {
			if newData == refData {
				fmt.Printf("  = %s\n", name)
			}
		}
	}
}
