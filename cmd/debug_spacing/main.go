package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()

	// Check line spacing values
	fmt.Printf("=== Line Spacing Distribution ===\n")
	lsDist := make(map[int32]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				if para.LineSpacing != 0 {
					lsDist[para.LineSpacing]++
				}
			}
		}
	}
	for ls, count := range lsDist {
		if ls > 0 {
			fmt.Printf("  LineSpacing=%d (percentage=%d%%): %d paragraphs\n", ls, ls, count)
		} else {
			fmt.Printf("  LineSpacing=%d (absolute=%d centipoints = %.1fpt): %d paragraphs\n", ls, -ls, float64(-ls)/100.0, count)
		}
	}

	// Check space before values
	fmt.Printf("\n=== Space Before Distribution ===\n")
	sbDist := make(map[int32]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				if para.SpaceBefore != 0 {
					sbDist[para.SpaceBefore]++
				}
			}
		}
	}
	for sb, count := range sbDist {
		if sb > 0 {
			fmt.Printf("  SpaceBefore=%d (percentage=%d%%): %d paragraphs\n", sb, sb, count)
		} else {
			fmt.Printf("  SpaceBefore=%d (absolute=%d centipoints = %.1fpt): %d paragraphs\n", sb, -sb, float64(-sb)/100.0, count)
		}
	}

	// Check space after values
	fmt.Printf("\n=== Space After Distribution ===\n")
	saDist := make(map[int32]int)
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				if para.SpaceAfter != 0 {
					saDist[para.SpaceAfter]++
				}
			}
		}
	}
	for sa, count := range saDist {
		if sa > 0 {
			fmt.Printf("  SpaceAfter=%d (percentage=%d%%): %d paragraphs\n", sa, sa, count)
		} else {
			fmt.Printf("  SpaceAfter=%d (absolute=%d centipoints = %.1fpt): %d paragraphs\n", sa, -sa, float64(-sa)/100.0, count)
		}
	}

	// Check bullet distribution
	fmt.Printf("\n=== Bullet Distribution ===\n")
	bulletDist := make(map[string]int)
	noBulletCount := 0
	hasBulletNoBulletChar := 0
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				if para.HasBullet {
					if para.BulletChar != "" {
						bulletDist[para.BulletChar]++
					} else {
						hasBulletNoBulletChar++
					}
				} else {
					noBulletCount++
				}
			}
		}
	}
	fmt.Printf("  No bullet: %d\n", noBulletCount)
	fmt.Printf("  HasBullet but no char: %d\n", hasBulletNoBulletChar)
	for char, count := range bulletDist {
		fmt.Printf("  BulletChar='%s': %d\n", char, count)
	}

	// Check indent/margin distribution
	fmt.Printf("\n=== Indent/Margin Distribution ===\n")
	hasMargin := 0
	hasIndent := 0
	for _, s := range slides {
		for _, sh := range s.GetShapes() {
			for _, para := range sh.Paragraphs {
				if para.LeftMargin != 0 {
					hasMargin++
				}
				if para.Indent != 0 {
					hasIndent++
				}
			}
		}
	}
	fmt.Printf("  Paragraphs with LeftMargin: %d\n", hasMargin)
	fmt.Printf("  Paragraphs with Indent: %d\n", hasIndent)
}
