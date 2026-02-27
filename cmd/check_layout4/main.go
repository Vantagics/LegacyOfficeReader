package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strings"
)

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	// Print layout 4 XML (used by 56 slides)
	for _, f := range zr.File {
		if f.Name == "ppt/slideLayouts/slideLayout4.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)
			fmt.Printf("Layout 4 size: %d bytes\n", len(data))
			fmt.Printf("Has bg: %v\n", strings.Contains(content, "<p:bg>"))
			fmt.Printf("Has blipFill: %v\n", strings.Contains(content, "blipFill"))
			fmt.Printf("Has cxnSp: %v\n", strings.Contains(content, "cxnSp"))
			fmt.Printf("Has p:sp: %v\n", strings.Contains(content, "<p:sp>"))
			fmt.Printf("Has p:pic: %v\n", strings.Contains(content, "<p:pic>"))
			fmt.Printf("showMasterSp: %v\n", strings.Contains(content, "showMasterSp"))
			fmt.Println("\nFull XML:")
			fmt.Println(content)
		}
	}
}
