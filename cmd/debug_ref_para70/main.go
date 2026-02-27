package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	// Read para 70 from reference docx
	refDoc := readZipEntry("testfie/test.docx", "word/document.xml")
	ourDoc := readZipEntry("testfie/test_new8.docx", "word/document.xml")

	refParas := splitParas(refDoc)
	ourParas := splitParas(ourDoc)

	fmt.Printf("Reference para 70 length: %d bytes\n", len(refParas[70]))
	fmt.Printf("Our para 70 length: %d bytes\n", len(ourParas[70]))

	// Extract text content from para 70 in reference
	refText := extractText(refParas[70])
	ourText := extractText(ourParas[70])

	refRunes := []rune(refText)
	ourRunes := []rune(ourText)

	fmt.Printf("\nReference text length: %d chars\n", len(refRunes))
	fmt.Printf("Our text length: %d chars\n", len(ourRunes))

	// Check if reference text is duplicated
	if len(refRunes) > 100 {
		half := len(refRunes) / 2
		first := string(refRunes[:60])
		atHalf := string(refRunes[half : half+60])
		fmt.Printf("\nRef first 60: %s\n", first)
		fmt.Printf("Ref at half:  %s\n", atHalf)
		if first == atHalf {
			fmt.Println("*** REFERENCE HAS DUPLICATED TEXT ***")
		}
	}

	// Show the runs in reference para 70
	fmt.Println("\n=== Reference para 70 runs ===")
	showRuns(refParas[70])
	fmt.Println("\n=== Our para 70 runs ===")
	showRuns(ourParas[70])
}

func showRuns(xml string) {
	rest := xml
	runNum := 0
	for {
		idx := strings.Index(rest, "<w:r>")
		if idx < 0 {
			idx = strings.Index(rest, "<w:r ")
		}
		if idx < 0 {
			break
		}
		endIdx := strings.Index(rest[idx:], "</w:r>")
		if endIdx < 0 {
			break
		}
		run := rest[idx : idx+endIdx+6]
		text := extractText(run)
		runes := []rune(text)
		runNum++
		if len(runes) > 80 {
			fmt.Printf("  Run %d: %d chars: %s...\n", runNum, len(runes), string(runes[:80]))
		} else {
			fmt.Printf("  Run %d: %d chars: %s\n", runNum, len(runes), text)
		}
		rest = rest[idx+endIdx+6:]
	}
}

func extractText(xml string) string {
	var sb strings.Builder
	rest := xml
	for {
		idx := strings.Index(rest, "<w:t")
		if idx < 0 {
			break
		}
		// Find >
		gt := strings.Index(rest[idx:], ">")
		if gt < 0 {
			break
		}
		start := idx + gt + 1
		end := strings.Index(rest[start:], "</w:t>")
		if end < 0 {
			break
		}
		sb.WriteString(rest[start : start+end])
		rest = rest[start+end+6:]
	}
	return sb.String()
}

func readZipEntry(zipPath, entry string) string {
	r, err := zip.OpenReader(zipPath)
	if err != nil {
		panic(err)
	}
	defer r.Close()
	for _, f := range r.File {
		if f.Name == entry {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			return string(data)
		}
	}
	panic("not found: " + entry)
}

func splitParas(xml string) []string {
	var result []string
	rest := xml
	for {
		pIdx := strings.Index(rest, "<w:p ")
		p2Idx := strings.Index(rest, "<w:p>")
		tIdx := strings.Index(rest, "<w:tbl")

		minIdx := -1
		if pIdx >= 0 {
			minIdx = pIdx
		}
		if p2Idx >= 0 && (minIdx < 0 || p2Idx < minIdx) {
			minIdx = p2Idx
		}
		if tIdx >= 0 && (minIdx < 0 || tIdx < minIdx) {
			minIdx = tIdx
		}
		if minIdx < 0 {
			break
		}
		if minIdx == tIdx {
			endTag := "</w:tbl>"
			endIdx := strings.Index(rest[minIdx:], endTag)
			if endIdx < 0 {
				result = append(result, rest[minIdx:])
				break
			}
			result = append(result, rest[minIdx:minIdx+endIdx+len(endTag)])
			rest = rest[minIdx+endIdx+len(endTag):]
		} else {
			endTag := "</w:p>"
			endIdx := strings.Index(rest[minIdx:], endTag)
			if endIdx < 0 {
				result = append(result, rest[minIdx:])
				break
			}
			result = append(result, rest[minIdx:minIdx+endIdx+len(endTag)])
			rest = rest[minIdx+endIdx+len(endTag):]
		}
	}
	return result
}
