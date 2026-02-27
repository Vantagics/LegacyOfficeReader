package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		panic(err)
	}
	defer f.Close()
	p, err := ppt.OpenReader(f)
	if err != nil {
		panic(err)
	}

	slides := p.GetSlides()
	s := slides[40] // slide 41
	shapes := s.GetShapes()

	// Check shapes around the "数据安全管控平台（DSCP）" text
	for si, sh := range shapes {
		// Check shapes near y=2268537
		if sh.Top >= 2200000 && sh.Top <= 2400000 {
			text := ""
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					text += run.Text
				}
			}
			kind := "SHAPE"
			if sh.IsImage {
				kind = fmt.Sprintf("IMAGE(idx=%d)", sh.ImageIdx)
			} else if sh.IsText {
				kind = "TEXT"
			}
			fmt.Printf("[%d] type=%d %s pos=(%d,%d) sz=(%d,%d) fill=%q noFill=%v text=%q\n",
				si, sh.ShapeType, kind, sh.Left, sh.Top, sh.Width, sh.Height,
				sh.FillColor, sh.NoFill, text)
		}
	}

	// Also check shape[28] which is the title
	fmt.Println("\n=== Title shape ===")
	sh := shapes[28]
	text := ""
	for _, para := range sh.Paragraphs {
		for _, run := range para.Runs {
			text += run.Text
			fmt.Printf("  Run: color=%q colorRaw=0x%08X text=%q\n", run.Color, run.ColorRaw, run.Text)
		}
	}
	fmt.Printf("[28] type=%d pos=(%d,%d) sz=(%d,%d) fill=%q noFill=%v text=%q\n",
		sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.NoFill, text)

	// Check shape[2] - DSCP text
	fmt.Println("\n=== DSCP shape ===")
	sh2 := shapes[2]
	for _, para := range sh2.Paragraphs {
		for _, run := range para.Runs {
			fmt.Printf("  Run: color=%q colorRaw=0x%08X text=%q\n", run.Color, run.ColorRaw, run.Text)
		}
	}
	fmt.Printf("[2] type=%d pos=(%d,%d) sz=(%d,%d) fill=%q noFill=%v fillRaw=0x%08X\n",
		sh2.ShapeType, sh2.Left, sh2.Top, sh2.Width, sh2.Height, sh2.FillColor, sh2.NoFill, sh2.FillColorRaw)

	// Check what's behind shape[2] at pos=(1820862,2268537)
	fmt.Println("\n=== Shapes overlapping DSCP position ===")
	for si, sh := range shapes {
		if si == 2 {
			continue
		}
		shRight := int64(sh.Left) + int64(sh.Width)
		shBottom := int64(sh.Top) + int64(sh.Height)
		if int64(sh.Left) <= 1820862 && shRight >= 1820862+6421437 &&
			int64(sh.Top) <= 2268537 && shBottom >= 2268537+460375 {
			text := ""
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					text += run.Text
				}
			}
			fmt.Printf("[%d] type=%d pos=(%d,%d) sz=(%d,%d) fill=%q noFill=%v text=%q\n",
				si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.NoFill, text)
		}
	}
}
