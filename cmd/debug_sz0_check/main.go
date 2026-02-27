package main

import (
	"archive/zip"
	"fmt"
	"os"
	"strings"
)

func main() {
	f, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	totalSz0 := 0
	for _, file := range f.File {
		if !strings.HasPrefix(file.Name, "ppt/slides/slide") || !strings.HasSuffix(file.Name, ".xml") {
			continue
		}
		rc, _ := file.Open()
		buf := make([]byte, file.UncompressedSize64)
		n, _ := rc.Read(buf)
		rc.Close()
		content := string(buf[:n])

		count := strings.Count(content, `sz="0"`)
		if count > 0 {
			totalSz0 += count
			fmt.Printf("%s: %d occurrences of sz=\"0\"\n", file.Name, count)
		}

		// Also check for unresolved scheme colors (colors starting with 00 that look suspicious)
		// Check for colors that might be scheme refs leaked through
		if strings.Contains(content, `val="000008"`) || strings.Contains(content, `val="0000FE"`) {
			fmt.Printf("%s: POSSIBLE UNRESOLVED SCHEME COLOR\n", file.Name)
		}
	}

	// Also check layouts
	for _, file := range f.File {
		if !strings.HasPrefix(file.Name, "ppt/slideLayouts/") || !strings.HasSuffix(file.Name, ".xml") {
			continue
		}
		rc, _ := file.Open()
		buf := make([]byte, file.UncompressedSize64)
		n, _ := rc.Read(buf)
		rc.Close()
		content := string(buf[:n])

		count := strings.Count(content, `sz="0"`)
		if count > 0 {
			totalSz0 += count
			fmt.Printf("%s: %d occurrences of sz=\"0\"\n", file.Name, count)
		}
	}

	if totalSz0 == 0 {
		fmt.Println("No sz=\"0\" found in any slide or layout XML")
	} else {
		fmt.Printf("\nTotal sz=\"0\" occurrences: %d\n", totalSz0)
	}
}
