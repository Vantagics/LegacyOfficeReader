package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/doc"
)

func main() {
	f, err := os.Open("testfie/test.doc")
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	document, err := doc.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "FAIL: %v\n", err)
		os.Exit(1)
	}

	fc := document.GetFormattedContent()
	if fc == nil {
		fmt.Println("No formatted content")
		os.Exit(1)
	}

	// Check body text paragraphs (P[136]-P[151]) for formatting details
	fmt.Println("=== Body text formatting (P[136]-P[151]) ===")
	for i := 136; i <= 151 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 60 {
			text = text[:60] + "..."
		}
		align := "L"
		switch p.Props.Alignment {
		case 1:
			align = "C"
		case 2:
			align = "R"
		case 3:
			align = "J"
		}
		if p.Props.AlignmentSet {
			align += "*"
		}
		fmt.Printf("P[%d] align=%s indent=%d/%d/%d sp=%d/%d line=%d/%d",
			i, align, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule)
		if p.HeadingLevel > 0 {
			fmt.Printf(" H%d", p.HeadingLevel)
		}
		if p.HasPageBreak {
			fmt.Printf(" PGBRK")
		}
		if p.IsListItem {
			fmt.Printf(" LIST")
		}
		fmt.Printf(": %q\n", text)
	}

	// Check the "产品概述" section (P[152]-P[155])
	fmt.Println("\n=== 产品概述 section (P[152]-P[158]) ===")
	for i := 152; i <= 158 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		align := "L"
		switch p.Props.Alignment {
		case 1:
			align = "C"
		case 2:
			align = "R"
		case 3:
			align = "J"
		}
		if p.Props.AlignmentSet {
			align += "*"
		}
		fmt.Printf("P[%d] align=%s indent=%d/%d/%d sp=%d/%d line=%d/%d",
			i, align, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule)
		if p.HeadingLevel > 0 {
			fmt.Printf(" H%d", p.HeadingLevel)
		}
		if p.HasPageBreak {
			fmt.Printf(" PGBRK")
		}
		fmt.Printf(": %q\n", text)
		for j, r := range p.Runs {
			rText := r.Text
			if len(rText) > 60 {
				rText = rText[:60] + "..."
			}
			fmt.Printf("  Run[%d]: font=%q sz=%d bold=%v italic=%v color=%q: %q\n",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Italic, r.Props.Color, rText)
		}
	}

	// Check the "产品优势与特点" section (P[172]-P[194])
	fmt.Println("\n=== 产品优势与特点 section (P[172]-P[194]) ===")
	for i := 172; i <= 194 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		align := "L"
		switch p.Props.Alignment {
		case 1:
			align = "C"
		case 2:
			align = "R"
		case 3:
			align = "J"
		}
		if p.Props.AlignmentSet {
			align += "*"
		}
		fmt.Printf("P[%d] align=%s indent=%d/%d/%d sp=%d/%d line=%d/%d",
			i, align, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule)
		if p.HeadingLevel > 0 {
			fmt.Printf(" H%d", p.HeadingLevel)
		}
		if p.IsListItem {
			fmt.Printf(" LIST(type=%d lvl=%d)", p.ListType, p.ListLevel)
		}
		fmt.Printf(": %q\n", text)
	}

	// Check title page paragraphs formatting
	fmt.Println("\n=== Title page formatting (P[33]-P[41]) ===")
	for i := 33; i <= 41 && i < len(fc.Paragraphs); i++ {
		p := fc.Paragraphs[i]
		text := ""
		for _, r := range p.Runs {
			text += r.Text
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		align := "L"
		switch p.Props.Alignment {
		case 1:
			align = "C"
		case 2:
			align = "R"
		case 3:
			align = "J"
		}
		if p.Props.AlignmentSet {
			align += "*"
		}
		fmt.Printf("P[%d] align=%s indent=%d/%d/%d sp=%d/%d line=%d/%d: %q\n",
			i, align, p.Props.IndentLeft, p.Props.IndentRight, p.Props.IndentFirst,
			p.Props.SpaceBefore, p.Props.SpaceAfter, p.Props.LineSpacing, p.Props.LineRule, text)
		for j, r := range p.Runs {
			rText := r.Text
			if len(rText) > 60 {
				rText = rText[:60] + "..."
			}
			fmt.Printf("  Run[%d]: font=%q sz=%d bold=%v color=%q: %q\n",
				j, r.Props.FontName, r.Props.FontSize, r.Props.Bold, r.Props.Color, rText)
		}
	}
}
