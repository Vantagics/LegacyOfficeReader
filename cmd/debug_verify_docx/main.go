package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Check section properties - should have section break for cover page
			sectPrCount := strings.Count(content, "<w:sectPr")
			fmt.Printf("Section properties: %d (should be 2 - one for cover page section, one final)\n", sectPrCount)

			// Check if cover page section has proper section break
			if strings.Contains(content, `<w:type w:val="nextPage"`) {
				fmt.Println("OK: Has nextPage section break for cover page")
			} else {
				fmt.Println("ISSUE: Missing nextPage section break for cover page")
			}

			// Check for evenAndOddHeaders in settings
			fmt.Println()

			// Check text content integrity - look for key Chinese text
			keyTexts := []string{
				"奇安信天眼威胁监测与分析系统",
				"创建时间",
				"修改时间",
				"目 录",
				"引言",
				"产品概述",
				"产品组成与架构",
				"威胁情报",
				"分析平台",
				"流量传感器",
				"文件威胁鉴定器",
				"其他周边组件",
				"产品优势与特点",
				"产品价值",
				"典型部署",
				"高级威胁检测及回溯方案",
				"部署拓扑图",
				"本地威胁发现方案",
				"文件威胁检测方案",
			}
			fmt.Println("=== Key text presence check ===")
			for _, t := range keyTexts {
				if strings.Contains(content, t) {
					fmt.Printf("  OK: %s\n", t)
				} else {
					fmt.Printf("  MISSING: %s\n", t)
				}
			}

			// Check table structure
			tblCount := strings.Count(content, "<w:tbl>")
			trCount := strings.Count(content, "<w:tr>")
			tcCount := strings.Count(content, "<w:tc>")
			fmt.Printf("\nTable: %d tables, %d rows, %d cells\n", tblCount, trCount, tcCount)

			// Check image count
			inlineImgCount := strings.Count(content, "<wp:inline")
			anchorImgCount := strings.Count(content, "<wp:anchor")
			fmt.Printf("Images: %d inline, %d anchored\n", inlineImgCount, anchorImgCount)

			// Check heading count
			for i := 1; i <= 3; i++ {
				hCount := strings.Count(content, fmt.Sprintf(`w:val="Heading%d"`, i))
				fmt.Printf("Heading%d: %d\n", i, hCount)
			}

			// Check list numbering
			numIdCount := strings.Count(content, "<w:numId")
			fmt.Printf("List items (numId): %d\n", numIdCount)

			// Check page breaks
			pbCount := strings.Count(content, `w:type="page"`)
			fmt.Printf("Page breaks: %d\n", pbCount)

			// Check TOC
			tocFieldCount := strings.Count(content, "TOC \\o")
			fmt.Printf("TOC fields: %d\n", tocFieldCount)
		}
	}

	// Check settings.xml for evenAndOddHeaders
	for _, f := range r.File {
		if f.Name == "word/settings.xml" {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			if strings.Contains(content, "evenAndOddHeaders") {
				fmt.Println("\nOK: settings.xml has evenAndOddHeaders")
			} else {
				fmt.Println("\nISSUE: settings.xml missing evenAndOddHeaders (needed for even page headers)")
			}
		}
	}
}
