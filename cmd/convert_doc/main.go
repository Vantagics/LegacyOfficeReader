package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/convert/docconv"
)

func main() {
	outPath := "testfie/test.docx"
	if len(os.Args) > 1 {
		outPath = os.Args[1]
	}
	if err := docconv.ConvertFile("testfie/test.doc", outPath); err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("OK: testfie/test.doc → %s\n", outPath)
}
