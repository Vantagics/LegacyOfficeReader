package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
)

type shapeInfo struct {
	x, y, w, h int
	fillColor  string
	hasNoFill  bool
	whiteTexts []string
}

func main() {
	zr, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	totalInvisible := 0
	for si := 1; si <= 71; si++ {
		name := fmt.Sprintf("ppt/slides/slide%d.xml", si)
		for _, f := range zr.File {
			if f.Name != name {
				continue
			}
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			content := string(data)

			shapes := parseShapes(content)

			for i, sh := range shapes {
				if !sh.hasNoFill || len(sh.whiteTexts) == 0 {
					continue
				}

				// Check if any earlier shape with a colored fill covers this shape
				centerX := sh.x + sh.w/2
				centerY := sh.y + sh.h/2
				hasBgShape := false
				for j := 0; j < i; j++ {
					bg := shapes[j]
					if bg.hasNoFill || bg.fillColor == "" {
						continue
					}
					// Check overlap
					if centerX >= bg.x && centerX <= bg.x+bg.w &&
						centerY >= bg.y && centerY <= bg.y+bg.h {
						hasBgShape = true
						break
					}
				}

				// Check if in title area (y < 1212850)
				inTitleArea := centerY < 1212850

				if !hasBgShape && !inTitleArea {
					allText := strings.Join(sh.whiteTexts, " | ")
					if len(allText) > 80 {
						allText = allText[:80] + "..."
					}
					fmt.Printf("INVISIBLE Slide %d shape %d: y=%d text=%s\n", si, i, sh.y, allText)
					totalInvisible++
				}
			}
		}
	}
	fmt.Printf("\nTotal truly invisible: %d\n", totalInvisible)
}

func parseShapes(content string) []shapeInfo {
	var shapes []shapeInfo

	// Split by </p:sp> to get individual shapes
	parts := strings.Split(content, "</p:sp>")
	for _, part := range parts {
		spPrIdx := strings.Index(part, "<p:spPr>")
		if spPrIdx < 0 {
			continue
		}
		spPrEnd := strings.Index(part[spPrIdx:], "</p:spPr>")
		if spPrEnd < 0 {
			continue
		}
		spPr := part[spPrIdx : spPrIdx+spPrEnd]

		// Get position
		offIdx := strings.Index(part, `<a:off x="`)
		if offIdx < 0 {
			continue
		}
		x := extractInt(part, offIdx+10)
		yIdx := strings.Index(part[offIdx:], `y="`)
		y := extractInt(part[offIdx:], yIdx+3)

		extIdx := strings.Index(part, `<a:ext cx="`)
		cx, cy := 0, 0
		if extIdx >= 0 {
			cx = extractInt(part, extIdx+11)
			cyIdx := strings.Index(part[extIdx:], `cy="`)
			cy = extractInt(part[extIdx:], cyIdx+4)
		}

		// Get fill from spPr (before <a:ln)
		lnIdx := strings.Index(spPr, "<a:ln")
		fillSection := spPr
		if lnIdx > 0 {
			fillSection = spPr[:lnIdx]
		}

		hasNoFill := strings.Contains(fillSection, "<a:noFill/>")
		fillColor := ""
		if strings.Contains(fillSection, "<a:solidFill>") {
			ci := strings.Index(fillSection, `srgbClr val="`)
			if ci >= 0 {
				ce := strings.Index(fillSection[ci+13:], `"`)
				fillColor = fillSection[ci+13 : ci+13+ce]
			}
		}

		// Get white texts
		var whiteTexts []string
		txBodyIdx := strings.Index(part, "<p:txBody>")
		if txBodyIdx >= 0 {
			txBody := part[txBodyIdx:]
			idx := 0
			for {
				ri := strings.Index(txBody[idx:], "<a:r>")
				if ri < 0 {
					break
				}
				idx += ri
				re := strings.Index(txBody[idx:], "</a:r>")
				if re < 0 {
					break
				}
				run := txBody[idx : idx+re]
				if strings.Contains(run, `val="FFFFFF"`) {
					text := extractText(run)
					if strings.TrimSpace(text) != "" {
						whiteTexts = append(whiteTexts, text)
					}
				}
				idx += re + 6
			}
		}

		shapes = append(shapes, shapeInfo{
			x: x, y: y, w: cx, h: cy,
			fillColor: fillColor, hasNoFill: hasNoFill,
			whiteTexts: whiteTexts,
		})
	}

	// Also parse pics (they can provide background)
	picParts := strings.Split(content, "</p:pic>")
	for _, part := range picParts {
		offIdx := strings.Index(part, `<a:off x="`)
		if offIdx < 0 {
			continue
		}
		x := extractInt(part, offIdx+10)
		yIdx := strings.Index(part[offIdx:], `y="`)
		y := extractInt(part[offIdx:], yIdx+3)
		extIdx := strings.Index(part, `<a:ext cx="`)
		cx, cy := 0, 0
		if extIdx >= 0 {
			cx = extractInt(part, extIdx+11)
			cyIdx := strings.Index(part[extIdx:], `cy="`)
			cy = extractInt(part[extIdx:], cyIdx+4)
		}
		// Images count as colored backgrounds
		shapes = append(shapes, shapeInfo{
			x: x, y: y, w: cx, h: cy,
			fillColor: "IMAGE", hasNoFill: false,
		})
	}

	return shapes
}

func extractInt(s string, start int) int {
	end := strings.IndexAny(s[start:], `"/ >`)
	if end < 0 {
		return 0
	}
	v, _ := strconv.Atoi(s[start : start+end])
	return v
}

func extractText(run string) string {
	ti := strings.Index(run, "<a:t>")
	if ti >= 0 {
		te := strings.Index(run[ti+5:], "</a:t>")
		if te >= 0 {
			return run[ti+5 : ti+5+te]
		}
	}
	ti = strings.Index(run, `<a:t xml:space="preserve">`)
	if ti >= 0 {
		te := strings.Index(run[ti+25:], "</a:t>")
		if te >= 0 {
			return run[ti+25 : ti+25+te]
		}
	}
	return ""
}
