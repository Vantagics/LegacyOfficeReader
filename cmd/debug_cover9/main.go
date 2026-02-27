package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "word/document.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			// Check if the title text is correct
			if strings.Contains(content, "奇安信天眼威胁监测与分析系统") {
				fmt.Println("OK: Title text found correctly in UTF-8")
			} else {
				fmt.Println("WARN: Title text not found - possible encoding issue")
				// Search for partial match
				if strings.Contains(content, "奇安信") {
					fmt.Println("  Found partial: 奇安信")
				}
			}

			// Check dates
			if strings.Contains(content, "创建时间") {
				fmt.Println("OK: Creation date text found")
			}
			if strings.Contains(content, "修改时间") {
				fmt.Println("OK: Modification date text found")
			}

			// Check key content
			checks := []string{
				"版权声明",
				"产品概述",
				"产品组成与架构",
				"典型部署",
				"产品价值",
			}
			for _, check := range checks {
				if strings.Contains(content, check) {
					fmt.Printf("OK: Found %q\n", check)
				} else {
					fmt.Printf("WARN: Missing %q\n", check)
				}
			}

			// Check header files
			fmt.Println("\n=== ZIP Contents ===")
			for _, zf := range r.File {
				if strings.HasPrefix(zf.Name, "word/header") || strings.HasPrefix(zf.Name, "word/footer") {
					rc2, _ := zf.Open()
					d2, _ := io.ReadAll(rc2)
					rc2.Close()
					fmt.Printf("%s: %d bytes\n", zf.Name, len(d2))
					// Check for images in headers
					if strings.Contains(string(d2), "r:embed") {
						fmt.Printf("  Has embedded image\n")
					}
				}
			}
		}
	}
}
