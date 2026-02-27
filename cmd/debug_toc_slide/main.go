package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Parse error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	// Find the TOC slide - look for slides with "背景与挑战" text
	for i, s := range slides {
		shapes := s.GetShapes()
		for _, sh := range shapes {
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					if len(run.Text) > 0 && (contains(run.Text, "背景与挑战") || contains(run.Text, "目") || contains(run.Text, "CONTENTS")) {
						fmt.Printf("Slide %d: found '%s'\n", i+1, run.Text)
					}
				}
			}
		}
	}
}

func contains(s, sub string) bool {
	for i := 0; i <= len(s)-len(sub); i++ {
		if s[i:i+len(sub)] == sub {
			return true
		}
	}
	return false
}
