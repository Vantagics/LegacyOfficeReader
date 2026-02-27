package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	slides := p.GetSlides()
	slide := slides[70]
	shapes := slide.GetShapes()

	// Simulate estimateDefaultFontSize for Shape[36]
	sh := shapes[36]
	fmt.Printf("Shape[36]: size=(%d,%d)\n", sh.Width, sh.Height)

	heightEMU := int64(sh.Height)
	widthEMU := int64(sh.Width)
	marginTop := int64(45720)
	marginBottom := int64(45720)
	marginLeft := int64(91440)
	marginRight := int64(91440)
	availHeight := heightEMU - marginTop - marginBottom
	availWidth := widthEMU - marginLeft - marginRight

	fmt.Printf("availHeight=%d availWidth=%d\n", availHeight, availWidth)

	// Check line spacing
	for pi, para := range sh.Paragraphs {
		fmt.Printf("Para[%d] lineSpacing=%d\n", pi, para.LineSpacing)
	}

	// Determine max line spacing
	lineSpacingMult := 1.0
	for _, para := range sh.Paragraphs {
		if para.LineSpacing > 0 {
			mult := float64(para.LineSpacing) / 100.0
			if mult > lineSpacingMult {
				lineSpacingMult = mult
			}
		}
	}
	fmt.Printf("lineSpacingMult=%.1f\n", lineSpacingMult)

	// Count total chars
	totalChars := 0
	for _, para := range sh.Paragraphs {
		for _, run := range para.Runs {
			totalChars += len([]rune(run.Text))
		}
	}
	fmt.Printf("totalChars=%d paraCount=%d\n", totalChars, len(sh.Paragraphs))

	candidates := []uint16{2000, 1800, 1600, 1400, 1200, 1100, 1000, 900, 800, 700, 600}
	for _, szCp := range candidates {
		szPt := float64(szCp) / 100.0
		szEMU := szPt * 12700.0
		lineHeightEMU := szEMU * 1.2
		if lineSpacingMult > 1.0 {
			lineHeightEMU = szEMU * lineSpacingMult
		}
		charWidthEMU := szEMU * 0.85
		charsPerLine := float64(availWidth) / charWidthEMU

		visualLines := 0
		for _, para := range sh.Paragraphs {
			paraChars := 0
			for _, run := range para.Runs {
				for _, r := range run.Text {
					if r == '\v' || r == '\n' {
						visualLines++
						paraChars = 0
					} else {
						paraChars++
					}
				}
			}
			if paraChars > 0 {
				wrapLines := (paraChars + int(charsPerLine) - 1) / int(charsPerLine)
				visualLines += wrapLines
			} else {
				visualLines++
			}
		}

		totalHeight := float64(visualLines) * lineHeightEMU
		fits := totalHeight <= float64(availHeight)
		fmt.Printf("sz=%d: lineH=%.0f charsPerLine=%.1f visualLines=%d totalH=%.0f fits=%v\n",
			szCp, lineHeightEMU, charsPerLine, visualLines, totalHeight, fits)
		if fits {
			fmt.Printf("  -> Would pick %d\n", szCp)
			break
		}
	}
}
