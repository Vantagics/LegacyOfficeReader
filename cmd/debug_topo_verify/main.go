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

			// Find topology diagram captions
			captions := []string{"未知威胁检测及回溯方案实施拓扑图", "本地威胁发现方案实施拓扑图", "文件威胁检测方案实施拓扑图"}
			for _, cap := range captions {
				idx := strings.Index(content, cap)
				if idx < 0 { continue }
				
				// Find the paragraph before this one
				pStart := strings.LastIndex(content[:idx], "<w:p>")
				// Find the paragraph before that
				prevPEnd := strings.LastIndex(content[:pStart], "</w:p>")
				prevPStart := strings.LastIndex(content[:prevPEnd], "<w:p>")
				if prevPStart >= 0 {
					prevPara := content[prevPStart : prevPEnd+len("</w:p>")]
					hasDrawing := strings.Contains(prevPara, "<w:drawing>")
					if len(prevPara) > 200 { prevPara = prevPara[:200] + "..." }
					fmt.Printf("Caption: %s\n  Prev para has drawing: %v\n  Prev para: %s\n\n", cap, hasDrawing, prevPara)
				}
			}
		}
	}
}
