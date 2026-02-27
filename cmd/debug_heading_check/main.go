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

			// Find heading paragraphs
			headings := []string{"引言", "产品概述", "产品组成与架构", "威胁情报", "分析平台",
				"流量传感器", "文件威胁鉴定器", "其他周边组件", "产品优势与特点", "产品价值",
				"典型部署", "高级威胁检测及回溯方案", "部署拓扑图", "本地威胁发现方案", "文件威胁检测方案"}
			for _, h := range headings {
				idx := strings.Index(content, h)
				if idx < 0 { continue }
				// Find enclosing <w:p>
				pStart := strings.LastIndex(content[:idx], "<w:p>")
				pEnd := strings.Index(content[idx:], "</w:p>")
				if pStart < 0 || pEnd < 0 { continue }
				para := content[pStart : idx+pEnd+len("</w:p>")]
				
				// Check if it has heading style
				hasHeading := strings.Contains(para, "Heading")
				hasNumPr := strings.Contains(para, "numPr")
				
				if hasHeading {
					// Extract heading level
					hIdx := strings.Index(para, "Heading")
					level := ""
					if hIdx+7 < len(para) {
						level = string(para[hIdx+7])
					}
					fmt.Printf("H%s hasNum=%v: %s\n", level, hasNumPr, h)
				}
			}
		}
	}
}
