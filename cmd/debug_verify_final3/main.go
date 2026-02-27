package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"sort"
)

func main() {
	ref := "testfie/test.docx"
	our := "testfie/test_new9.docx"

	fmt.Printf("Comparing:\n  REF: %s\n  OUR: %s\n\n", ref, our)

	refFiles, _ := listZipFiles(ref)
	ourFiles, _ := listZipFiles(our)

	refSet := map[string]bool{}
	for _, f := range refFiles {
		refSet[f] = true
	}
	ourSet := map[string]bool{}
	for _, f := range ourFiles {
		ourSet[f] = true
	}

	for _, f := range refFiles {
		if !ourSet[f] {
			fmt.Printf("  MISSING in our: %s\n", f)
		}
	}
	for _, f := range ourFiles {
		if !refSet[f] {
			fmt.Printf("  EXTRA in our: %s\n", f)
		}
	}

	allSame := true
	for _, f := range refFiles {
		if !ourSet[f] {
			continue
		}
		refData, _ := readZipFile(ref, f)
		ourData, _ := readZipFile(our, f)
		if bytes.Equal(refData, ourData) {
			fmt.Printf("  SAME  %s (%d bytes)\n", f, len(refData))
		} else {
			fmt.Printf("  DIFF  %s (ref=%d, our=%d, delta=%d)\n", f, len(refData), len(ourData), len(ourData)-len(refData))
			allSame = false
		}
	}

	if allSame {
		fmt.Println("\nALL FILES IDENTICAL")
	} else {
		fmt.Println("\nFILES DIFFER")
	}
}

func readZipFile(zipPath, innerPath string) ([]byte, error) {
	r, err := zip.OpenReader(zipPath)
	if err != nil {
		return nil, err
	}
	defer r.Close()
	for _, f := range r.File {
		if f.Name == innerPath {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("not found: %s", innerPath)
}

func listZipFiles(zipPath string) ([]string, error) {
	r, err := zip.OpenReader(zipPath)
	if err != nil {
		return nil, err
	}
	defer r.Close()
	var names []string
	for _, f := range r.File {
		names = append(names, f.Name)
	}
	sort.Strings(names)
	return names, nil
}
