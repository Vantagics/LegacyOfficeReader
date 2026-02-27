package main

import (
	"fmt"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	presentation, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Println("Error:", err)
		return
	}

	slides := presentation.GetSlides()
	masters := presentation.GetMasters()

	// Check masters for watermark-like shapes
	fmt.Println("=== Master shapes analysis ===")
	for ref, m := range masters {
		fmt.Printf("\nMaster %d: bg=%v shapes=%d\n", ref, m.Background.HasBackground, len(m.Shapes))
		for i, sh := range m.Shapes {
			text := ""
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					text += r.Text
				}
			}
			desc := fmt.Sprintf("type=%d", sh.ShapeType)
			if sh.IsImage {
				desc = fmt.Sprintf("IMAGE(idx=%d)", sh.ImageIdx)
			}
			if sh.IsText {
				desc = "TEXT"
			}
			fillDesc := ""
			if sh.NoFill {
				fillDesc = "noFill"
			} else if sh.FillColor != "" {
				fillDesc = fmt.Sprintf("fill=%s", sh.FillColor)
				if sh.FillOpacity >= 0 && sh.FillOpacity < 65536 {
					fillDesc += fmt.Sprintf("@%d%%", sh.FillOpacity*100/65536)
				}
			}
			fmt.Printf("  [%d] %s pos=(%d,%d) size=(%d,%d) %s rot=%d", i, desc, sh.Left, sh.Top, sh.Width, sh.Height, fillDesc, sh.Rotation)
			if text != "" {
				if len(text) > 60 {
					text = text[:60] + "..."
				}
				fmt.Printf(" text=%q", text)
			}
			fmt.Println()
		}
	}

	// Check first few slides for watermark-like shapes (semi-transparent, rotated text)
	fmt.Println("\n=== Slide watermark analysis ===")
	for i, s := range slides {
		shapes := s.GetShapes()
		for j, sh := range shapes {
			// Look for watermark indicators: rotated text, semi-transparent, or specific text
			isWatermarkLike := false
			text := ""
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					text += r.Text
				}
			}

			if sh.Rotation != 0 && sh.IsText {
				isWatermarkLike = true
			}
			if sh.FillOpacity > 0 && sh.FillOpacity < 32768 { // < 50% opacity
				isWatermarkLike = true
			}
			if strings.Contains(strings.ToLower(text), "watermark") || strings.Contains(text, "水印") {
				isWatermarkLike = true
			}
			// Check for www.qianxin.com or similar footer/watermark text
			if strings.Contains(text, "www.") || strings.Contains(text, "qianxin") || strings.Contains(text, "奇安信") {
				if sh.Width > 5000000 || sh.Height < 500000 { // wide or short = likely footer/watermark
					isWatermarkLike = true
				}
			}

			if isWatermarkLike {
				fillDesc := ""
				if sh.NoFill {
					fillDesc = "noFill"
				} else if sh.FillColor != "" {
					fillDesc = fmt.Sprintf("fill=%s opacity=%d", sh.FillColor, sh.FillOpacity)
				}
				fmt.Printf("Slide %d shape %d: type=%d pos=(%d,%d) size=(%d,%d) rot=%d %s text=%q\n",
					i+1, j, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.Rotation, fillDesc, text)
			}
		}
	}

	// Check layout shapes (from masters used by slides)
	fmt.Println("\n=== Layout shapes that appear on slides ===")
	masterRefToSlideCount := make(map[uint32]int)
	for _, s := range slides {
		masterRefToSlideCount[s.GetMasterRef()]++
	}
	for ref, count := range masterRefToSlideCount {
		m, ok := masters[ref]
		if !ok {
			continue
		}
		fmt.Printf("\nMaster %d (used by %d slides): %d shapes\n", ref, count, len(m.Shapes))
		for i, sh := range m.Shapes {
			text := ""
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					text += r.Text
				}
			}
			desc := fmt.Sprintf("type=%d", sh.ShapeType)
			if sh.IsImage {
				desc = fmt.Sprintf("IMAGE(idx=%d)", sh.ImageIdx)
			} else if sh.IsText {
				desc = "TEXT"
			}
			fillDesc := ""
			if sh.NoFill {
				fillDesc = "noFill"
			} else if sh.FillColor != "" {
				fillDesc = fmt.Sprintf("fill=%s", sh.FillColor)
			}
			lineDesc := ""
			if sh.LineColor != "" {
				lineDesc = fmt.Sprintf("line=%s w=%d", sh.LineColor, sh.LineWidth)
			}
			fmt.Printf("  [%d] %s pos=(%d,%d) size=(%d,%d) %s %s", i, desc, sh.Left, sh.Top, sh.Width, sh.Height, fillDesc, lineDesc)
			if text != "" {
				if len(text) > 60 {
					text = text[:60] + "..."
				}
				fmt.Printf(" text=%q", text)
			}
			fmt.Println()
		}
	}
}
