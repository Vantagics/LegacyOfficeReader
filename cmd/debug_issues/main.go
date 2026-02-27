package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	pres, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := pres.GetSlides()
	masters := pres.GetMasters()

	// Issue 1: Check for white text on light fills (visibility issues)
	fmt.Println("=== Issue 1: White text on light fills ===")
	for i, s := range slides {
		shapes := s.GetShapes()
		ref := s.GetMasterRef()
		m, _ := masters[ref]
		scheme := m.ColorScheme

		for si, sh := range shapes {
			for pi, p := range sh.Paragraphs {
				for ri, r := range p.Runs {
					text := strings.TrimSpace(r.Text)
					if text == "" {
						continue
					}
					// White text on light fill
					if r.Color == "FFFFFF" && sh.FillColor != "" && !isDark(sh.FillColor) && !sh.NoFill {
						fmt.Printf("  Slide %d, Shape %d, Para %d, Run %d: WHITE text on LIGHT fill=%s\n", i+1, si, pi, ri, sh.FillColor)
						fmt.Printf("    Text: %s\n", truncate(text, 50))
						fmt.Printf("    ColorRaw=0x%08X, scheme=%v\n", r.ColorRaw, scheme)
					}
					// Scheme[0] color (usually FFFFFF) on light fill
					if len(scheme) > 0 && r.Color == scheme[0] && sh.FillColor != "" && !isDark(sh.FillColor) && !sh.NoFill {
						if r.Color != "FFFFFF" { // already caught above
							fmt.Printf("  Slide %d, Shape %d: scheme[0]=%s text on fill=%s\n", i+1, si, r.Color, sh.FillColor)
							fmt.Printf("    Text: %s\n", truncate(text, 50))
						}
					}
				}
			}
		}
	}

	// Issue 2: Check for sz=0 runs
	fmt.Println("\n=== Issue 2: Font size 0 runs ===")
	sz0Count := 0
	for i, s := range slides {
		shapes := s.GetShapes()
		for si, sh := range shapes {
			for pi, p := range sh.Paragraphs {
				for ri, r := range p.Runs {
					if r.FontSize == 0 && strings.TrimSpace(r.Text) != "" {
						sz0Count++
						if sz0Count <= 20 {
							fmt.Printf("  Slide %d, Shape %d, Para %d, Run %d: sz=0 font=%s bold=%v\n", i+1, si, pi, ri, r.FontName, r.Bold)
							fmt.Printf("    Text: %s\n", truncate(strings.TrimSpace(r.Text), 50))
							fmt.Printf("    Shape: type=%d pos=(%d,%d) sz=(%d,%d) fill=%s\n", sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor)
						}
					}
				}
			}
		}
	}
	fmt.Printf("  Total sz=0 runs: %d\n", sz0Count)

	// Issue 3: Check for text that might be invisible (same color as background)
	fmt.Println("\n=== Issue 3: Potentially invisible text ===")
	for i, s := range slides {
		shapes := s.GetShapes()
		ref := s.GetMasterRef()
		m, _ := masters[ref]

		for si, sh := range shapes {
			for pi, p := range sh.Paragraphs {
				for ri, r := range p.Runs {
					text := strings.TrimSpace(r.Text)
					if text == "" {
						continue
					}
					// White text on transparent shape with white master bg
					if r.Color == "FFFFFF" && sh.FillColor == "" && m.Background.FillColor == "FFFFFF" {
						// Check if there's a layout image behind
						hasLayoutImg := false
						for _, ms := range m.Shapes {
							if ms.IsImage && ms.ImageIdx >= 0 {
								// Check overlap
								if int64(sh.Left)+int64(sh.Width)/2 >= int64(ms.Left) &&
									int64(sh.Left)+int64(sh.Width)/2 <= int64(ms.Left)+int64(ms.Width) &&
									int64(sh.Top)+int64(sh.Height)/2 >= int64(ms.Top) &&
									int64(sh.Top)+int64(sh.Height)/2 <= int64(ms.Top)+int64(ms.Height) {
									hasLayoutImg = true
								}
							}
						}
						if !hasLayoutImg {
							// Check if title bg detection would save this
							hasConnector := false
							for _, ms := range m.Shapes {
								if isConnector(ms.ShapeType) && ms.Height == 0 {
									hasConnector = true
								}
							}
							if hasConnector && sh.Top+sh.Height/2 < 1300000 {
								continue // title bg gradient will handle this
							}
							fmt.Printf("  Slide %d, Shape %d, Para %d, Run %d: WHITE text on WHITE bg, no layout image\n", i+1, si, pi, ri)
							fmt.Printf("    Text: %s\n", truncate(text, 50))
							fmt.Printf("    ColorRaw=0x%08X bold=%v sz=%d\n", r.ColorRaw, r.Bold, r.FontSize)
						}
					}
				}
			}
		}
	}

	// Issue 4: Check slide 63 specifically for table text colors
	fmt.Println("\n=== Issue 4: Slide 63 table text analysis ===")
	if len(slides) >= 63 {
		s := slides[62]
		shapes := s.GetShapes()
		colorFillCombos := make(map[string]int)
		for _, sh := range shapes {
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					text := strings.TrimSpace(r.Text)
					if text == "" {
						continue
					}
					key := fmt.Sprintf("textColor=%s fill=%s noFill=%v raw=0x%08X", r.Color, sh.FillColor, sh.NoFill, r.ColorRaw)
					colorFillCombos[key]++
				}
			}
		}
		for combo, count := range colorFillCombos {
			fmt.Printf("  %s (count=%d)\n", combo, count)
		}
	}
}

func isDark(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	r := hexDigit(hex[0])*16 + hexDigit(hex[1])
	g := hexDigit(hex[2])*16 + hexDigit(hex[3])
	b := hexDigit(hex[4])*16 + hexDigit(hex[5])
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	return lum < 128
}

func hexDigit(c byte) int {
	switch {
	case c >= '0' && c <= '9':
		return int(c - '0')
	case c >= 'A' && c <= 'F':
		return int(c-'A') + 10
	case c >= 'a' && c <= 'f':
		return int(c-'a') + 10
	}
	return 0
}

func isConnector(shapeType uint16) bool {
	switch shapeType {
	case 20, 32, 33, 34, 35, 36, 37, 38, 39, 40:
		return true
	}
	return false
}

func truncate(s string, n int) string {
	runes := []rune(s)
	if len(runes) > n {
		return string(runes[:n]) + "..."
	}
	return s
}
