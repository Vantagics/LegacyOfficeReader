package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func readZipFile(f *zip.File) (string, error) {
	rc, err := f.Open()
	if err != nil {
		return "", err
	}
	defer rc.Close()
	data, err := io.ReadAll(rc)
	if err != nil {
		return "", err
	}
	return string(data), nil
}

func main() {
	r, err := zip.OpenReader("testfie/test.docx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	fmt.Println("=== DOCX Archive Contents ===")
	for _, f := range r.File {
		fmt.Printf("  %s (%d bytes)\n", f.Name, f.UncompressedSize64)
	}

	files := make(map[string]string)
	for _, f := range r.File {
		content, err := readZipFile(f)
		if err == nil {
			files[f.Name] = content
		}
	}

	if content, ok := files["word/document.xml"]; ok {
		fmt.Printf("\n=== document.xml (%d bytes) ===\n", len(content))
		pCount := strings.Count(content, "<w:p>") + strings.Count(content, "<w:p ")
		fmt.Printf("Paragraphs: %d\n", pCount)
		fmt.Printf("Tables: %d\n", strings.Count(content, "<w:tbl>"))
		fmt.Printf("Drawings: %d\n", strings.Count(content, "<w:drawing>"))
		fmt.Printf("Inline: %d, Anchor: %d\n", strings.Count(content, "<wp:inline"), strings.Count(content, "<wp:anchor"))
		fmt.Printf("TextBoxes: %d\n", strings.Count(content, "<wps:wsp>"))
		for lvl := 1; lvl <= 3; lvl++ {
			s := fmt.Sprintf(`<w:pStyle w:val="Heading%d"/>`, lvl)
			fmt.Printf("Heading%d: %d\n", lvl, strings.Count(content, s))
		}
		tocCount := 0
		for lvl := 1; lvl <= 3; lvl++ {
			s := fmt.Sprintf(`<w:pStyle w:val="TOC%d"/>`, lvl)
			tocCount += strings.Count(content, s)
		}
		fmt.Printf("TOC entries: %d\n", tocCount)
		fmt.Printf("TOC field: %v\n", strings.Contains(content, `TOC \o "1-3"`))
		fmt.Printf("Page breaks: %d\n", strings.Count(content, `<w:br w:type="page"/>`))
		sectCount := strings.Count(content, "<w:sectPr>") + strings.Count(content, "<w:sectPr ")
		fmt.Printf("SectPr: %d\n", sectCount)
		fmt.Printf("HeaderRef: %d\n", strings.Count(content, "headerReference"))
		fmt.Printf("FooterRef: %d\n", strings.Count(content, "footerReference"))
		fmt.Printf("A4: %v\n", strings.Contains(content, `w:w="11906" w:h="16838"`))
		fmt.Printf("NumPr: %d\n", strings.Count(content, "<w:numPr>"))
		fmt.Printf("Title: %v\n", strings.Contains(content, "奇安信天眼"))
		fmt.Printf("RevTable: %v\n", strings.Contains(content, "修订记录"))
		fmt.Printf("宋体: %d, 黑体: %d\n", strings.Count(content, "宋体"), strings.Count(content, "黑体"))
		for i := 1; i <= 11; i++ {
			rid := fmt.Sprintf(`r:embed="rImg%d"`, i)
			c := strings.Count(content, rid)
			if c > 0 {
				fmt.Printf("Image rImg%d: %d refs\n", i, c)
			}
		}
		// First paragraph
		idx := strings.Index(content, "<w:body>")
		if idx >= 0 {
			snippet := content[idx:]
			if len(snippet) > 600 {
				snippet = snippet[:600]
			}
			fmt.Printf("Body start:\n%s...\n", snippet)
		}
	}

	if content, ok := files["word/styles.xml"]; ok {
		fmt.Printf("\n=== styles.xml (%d bytes) ===\n", len(content))
		for lvl := 1; lvl <= 3; lvl++ {
			fmt.Printf("Heading%d: %v\n", lvl, strings.Contains(content, fmt.Sprintf(`w:styleId="Heading%d"`, lvl)))
		}
		for lvl := 1; lvl <= 3; lvl++ {
			fmt.Printf("TOC%d: %v\n", lvl, strings.Contains(content, fmt.Sprintf(`w:styleId="TOC%d"`, lvl)))
		}
		fmt.Printf("DotLeader: %v\n", strings.Contains(content, `w:leader="dot"`))
		fmt.Printf("Normal: %v\n", strings.Contains(content, `w:styleId="Normal"`))
		fmt.Printf("Footer: %v\n", strings.Contains(content, `w:styleId="Footer"`))
		fmt.Printf("Header: %v\n", strings.Contains(content, `w:styleId="Header"`))
		fmt.Printf("DocDefaults: %v\n", strings.Contains(content, "<w:docDefaults>"))
		fmt.Printf("Spacing0/240: %v\n", strings.Contains(content, `w:after="0" w:line="240"`))
	}

	if content, ok := files["word/numbering.xml"]; ok {
		fmt.Printf("\n=== numbering.xml (%d bytes) ===\n", len(content))
		fmt.Printf("AbstractNum: %d\n", strings.Count(content, "<w:abstractNum"))
		fmt.Printf("Num: %d\n", strings.Count(content, "<w:num "))
		fmt.Printf("Multilevel: %v\n", strings.Contains(content, `w:val="multilevel"`))
		fmt.Printf("Decimal: %v\n", strings.Contains(content, `w:val="decimal"`))
		fmt.Printf("Bullet: %v\n", strings.Contains(content, `w:val="bullet"`))
	}

	if content, ok := files["word/footer1.xml"]; ok {
		fmt.Printf("\n=== footer1.xml ===\n")
		fmt.Printf("PAGE: %v\n", strings.Contains(content, "PAGE"))
		fmt.Printf("NUMPAGES: %v\n", strings.Contains(content, "NUMPAGES"))
		fmt.Printf("FooterStyle: %v\n", strings.Contains(content, `w:val="Footer"`))
		fmt.Printf("Tabs: %v\n", strings.Contains(content, "<w:tab/>"))
		fmt.Printf("文档名称: %v\n", strings.Contains(content, "文档名称"))
		fmt.Printf("奇安信: %v\n", strings.Contains(content, "奇安信"))
	}

	if content, ok := files["word/_rels/document.xml.rels"]; ok {
		fmt.Printf("\n=== document.xml.rels ===\n")
		fmt.Printf("Styles: %v\n", strings.Contains(content, "styles.xml"))
		fmt.Printf("Numbering: %v\n", strings.Contains(content, "numbering.xml"))
		fmt.Printf("Header: %v\n", strings.Contains(content, "header"))
		fmt.Printf("Footer: %v\n", strings.Contains(content, "footer"))
		fmt.Printf("Images: %d\n", strings.Count(content, "relationships/image"))
	}

	if content, ok := files["[Content_Types].xml"]; ok {
		fmt.Printf("\n=== [Content_Types].xml ===\n")
		fmt.Printf("Document: %v\n", strings.Contains(content, "document.main+xml"))
		fmt.Printf("Styles: %v\n", strings.Contains(content, "styles+xml"))
		fmt.Printf("Numbering: %v\n", strings.Contains(content, "numbering+xml"))
		fmt.Printf("Header: %v\n", strings.Contains(content, "header+xml"))
		fmt.Printf("Footer: %v\n", strings.Contains(content, "footer+xml"))
	}
}
