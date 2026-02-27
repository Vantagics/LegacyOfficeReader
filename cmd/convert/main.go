package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/convert/docconv"
	"github.com/shakinm/xlsReader/convert/pptconv"
)

func main() {
	conversions := []struct {
		name   string
		input  string
		output string
		fn     func(string, string) error
	}{
		{"PPT → PPTX", "testfie/test.ppt", "testfie/test.pptx", pptconv.ConvertFile},
		{"DOC → DOCX", "testfie/test.doc", "testfie/test.docx", docconv.ConvertFile},
	}

	for _, c := range conversions {
		if err := c.fn(c.input, c.output); err != nil {
			fmt.Fprintf(os.Stderr, "FAIL %s: %v\n", c.name, err)
			continue
		}
		fmt.Printf("OK   %s: %s → %s\n", c.name, c.input, c.output)
	}
}
