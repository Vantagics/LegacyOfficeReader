package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
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

	newXML := readZip(r1, "word/document.xml")
	refXML := readZip(r2, "word/document.xml")

	paraRe := regexp.MustCompile(`<w:p[ >].*?</w:p>|<w:p/>`)
	newParas := paraRe.FindAllString(newXML, -1)
	refParas := paraRe.FindAllString(refXML, -1)

	fmt.Printf("New: %d paragraphs, Ref: %d paragraphs\n\n", len(newParas), len(refParas))

	maxLen := len(newParas)
	if len(refParas) > maxLen {
		maxLen = len(refParas)
	}

	diffCount := 0
	for i := 0; i < maxLen; i++ {
		newP := ""
		refP := ""
		if i < len(newParas) {
			newP = newParas[i]
		}
		if i < len(refParas) {
			refP = refParas[i]
		}

		if newP != refP {
			diffCount++
			textRe := regexp.MustCompile(`<w:t[^>]*>([^<]*)</w:t>`)

			newTexts := textRe.FindAllStringSubmatch(newP, -1)
			newText := ""
			for _, t := range newTexts {
				newText += t[1]
			}

			refTexts := textRe.FindAllStringSubmatch(refP, -1)
			refText := ""
			for _, t := range refTexts {
				refText += t[1]
			}

			if len(newText) > 40 {
				newText = newText[:40] + "..."
			}
			if len(refText) > 40 {
				refText = refText[:40] + "..."
			}

			fmt.Printf("DIFF at para %d:\n", i+1)
			fmt.Printf("  NEW text: %q\n", newText)
			fmt.Printf("  REF text: %q\n", refText)

			// Show PPR differences
			pprRe := regexp.MustCompile(`<w:pPr>(.*?)</w:pPr>`)
			newPPR := pprRe.FindStringSubmatch(newP)
			refPPR := pprRe.FindStringSubmatch(refP)
			newPPRStr := ""
			refPPRStr := ""
			if newPPR != nil {
				newPPRStr = newPPR[1]
			}
			if refPPR != nil {
				refPPRStr = refPPR[1]
			}
			if newPPRStr != refPPRStr {
				fmt.Printf("  NEW ppr: %s\n", newPPRStr)
				fmt.Printf("  REF ppr: %s\n", refPPRStr)
			}

			// Show run property differences (first run only)
			rprRe := regexp.MustCompile(`<w:rPr>(.*?)</w:rPr>`)
			newRPR := rprRe.FindStringSubmatch(newP)
			refRPR := rprRe.FindStringSubmatch(refP)
			newRPRStr := ""
			refRPRStr := ""
			if newRPR != nil {
				newRPRStr = newRPR[1]
			}
			if refRPR != nil {
				refRPRStr = refRPR[1]
			}
			if newRPRStr != refRPRStr {
				fmt.Printf("  NEW rpr: %s\n", newRPRStr)
				fmt.Printf("  REF rpr: %s\n", refRPRStr)
			}

			fmt.Println()
			if diffCount > 20 {
				fmt.Println("... (truncated, too many diffs)")
				break
			}
		}
	}

	if diffCount == 0 {
		fmt.Println("NO DIFFERENCES FOUND - files are identical!")
	} else {
		fmt.Printf("\nTotal differences: %d\n", diffCount)
	}
}

func readZip(r *zip.ReadCloser, name string) string {
	for _, f := range r.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			return string(data)
		}
	}
	return ""
}
