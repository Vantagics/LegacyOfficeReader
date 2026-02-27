package main

import (
	"fmt"
)

func main() {
	// Yellow box: height=503237, width=12192000
	// Text: "需要评估客户的EPS，每9000EPS增加一个集群节点" (~25 chars)
	heightEMU := int64(503237)
	widthEMU := int64(12192000)

	// Default margins
	marginTop := int64(45720)
	marginBottom := int64(45720)
	marginLeft := int64(91440)
	marginRight := int64(91440)

	availHeight := heightEMU - marginTop - marginBottom
	availWidth := widthEMU - marginLeft - marginRight

	fmt.Printf("Available: %d x %d EMU\n", availWidth, availHeight)

	for _, szCp := range []uint16{2400, 2000, 1800, 1600, 1400} {
		szPt := float64(szCp) / 100.0
		szEMU := szPt * 12700.0
		lineHeight := szEMU * 1.2

		charWidth := szEMU * 0.85
		charsPerLine := float64(availWidth) / charWidth
		totalChars := 25
		wrapLines := (totalChars + int(charsPerLine) - 1) / int(charsPerLine)
		totalHeight := float64(wrapLines) * lineHeight

		fits := totalHeight <= float64(availHeight)
		fmt.Printf("sz=%d: szEMU=%.0f lineH=%.0f charsPerLine=%.0f wrapLines=%d totalH=%.0f avail=%d fits=%v\n",
			szCp, szEMU, lineHeight, charsPerLine, wrapLines, totalHeight, availHeight, fits)
	}
}
