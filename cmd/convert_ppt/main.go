package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/convert/pptconv"
)

func main() {
	input := "testfie/test.ppt"
	output := "testfie/test.pptx"
	if len(os.Args) > 2 {
		input = os.Args[1]
		output = os.Args[2]
	}
	if err := pptconv.ConvertFile(input, output); err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("OK: %s → %s\n", input, output)
}
