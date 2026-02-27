package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"regexp"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// PPT source shapes
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	slides := p.GetSlides()
	slide := slides[70]
	shapes := slide.GetShapes()

	fmt.Printf("=== PPT Source: %d shapes ===\n", len(shapes))
	for i, sh := range shapes {
		text := ""
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				text += run.Text
			}
		}
		if len(text) > 50 {
			text = text[:50] + "..."
		}
		extra := ""
		if sh.IsImage {
			extra += fmt.Sprintf(" IMG[%d]", sh.ImageIdx)
		}
		if len(sh.GeoVertices) > 0 {
			extra += fmt.Sprintf(" GEO[v=%d,s=%d]", len(sh.GeoVertices), len(sh.GeoSegments))
		}
		if sh.FillColor != "" {
			extra += fmt.Sprintf(" fill=%s", sh.FillColor)
		}
		if sh.NoFill {
			extra += " noFill"
		}
		if sh.LineColor != "" {
			extra += fmt.Sprintf(" line=%s", sh.LineColor)
		}
		if sh.NoLine {
			extra += " noLine"
		}
		fmt.Printf("  [%d] type=%d pos=(%d,%d) sz=(%d,%d)%s text=%q\n",
			i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, extra, text)
	}

	// PPTX output
	fmt.Println("\n=== PPTX Output ===")
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "ppt/slides/slide71.xml" {
			rc, _ := f.Open()
			data, _ := io.ReadAll(rc)
			content := string(data)
			rc.Close()

			// Count shape types
			spCount := strings.Count(content, "<p:sp>")
			picCount := strings.Count(content, "<p:pic>")
			cxnCount := strings.Count(content, "<p:cxnSp>")
			grpCount := strings.Count(content, "<p:grpSp>")
			fmt.Printf("  Shapes: %d sp + %d pic + %d cxnSp + %d grpSp = %d total\n",
				spCount, picCount, cxnCount, grpCount,
				spCount+picCount+cxnCount+grpCount)

			// Extract each shape's details
			// Find all sp elements with their names and positions
			nameRe := regexp.MustCompile(`name="([^"]*)"`)
			offRe := regexp.MustCompile(`<a:off x="([^"]*)" y="([^"]*)"/>`)
			extRe := regexp.MustCompile(`<a:ext cx="([^"]*)" cy="([^"]*)"/>`)

			// Extract shapes by splitting on shape boundaries
			shapeStarts := []string{"<p:sp>", "<p:pic>", "<p:cxnSp>"}
			shapeEnds := []string{"</p:sp>", "</p:pic>", "</p:cxnSp>"}

			for si, startTag := range shapeStarts {
				endTag := shapeEnds[si]
				idx := 0
				count := 0
				for {
					pos := strings.Index(content[idx:], startTag)
					if pos < 0 {
						break
					}
					absStart := idx + pos
					endPos := strings.Index(content[absStart:], endTag)
					if endPos < 0 {
						break
					}
					shapeXML := content[absStart : absStart+endPos+len(endTag)]

					name := ""
					if m := nameRe.FindStringSubmatch(shapeXML); m != nil {
						name = m[1]
					}
					offX, offY := "?", "?"
					if m := offRe.FindStringSubmatch(shapeXML); m != nil {
						offX, offY = m[1], m[2]
					}
					extCx, extCy := "?", "?"
					if m := extRe.FindStringSubmatch(shapeXML); m != nil {
						extCx, extCy = m[1], m[2]
					}

					extra := ""
					if strings.Contains(shapeXML, "<a:custGeom>") {
						extra += " CUSTGEOM"
						if strings.Contains(shapeXML, "<a:cubicBezTo>") {
							bezCount := strings.Count(shapeXML, "<a:cubicBezTo>")
							extra += fmt.Sprintf("(bez=%d)", bezCount)
						}
					}
					if strings.Contains(shapeXML, "r:embed=") {
						embedRe := regexp.MustCompile(`r:embed="([^"]*)"`)
						if m := embedRe.FindStringSubmatch(shapeXML); m != nil {
							extra += fmt.Sprintf(" embed=%s", m[1])
						}
					}
					if strings.Contains(shapeXML, "<a:solidFill>") {
						fillRe := regexp.MustCompile(`<a:solidFill>\s*<a:srgbClr val="([^"]*)"`)
						if m := fillRe.FindStringSubmatch(shapeXML); m != nil {
							extra += fmt.Sprintf(" fill=%s", m[1])
						}
					}
					if strings.Contains(shapeXML, "<a:noFill/>") {
						extra += " noFill"
					}

					// Get text content
					textRe := regexp.MustCompile(`<a:t>([^<]*)</a:t>`)
					textMatches := textRe.FindAllStringSubmatch(shapeXML, -1)
					text := ""
					for _, tm := range textMatches {
						text += tm[1]
					}
					if len(text) > 40 {
						text = text[:40] + "..."
					}

					typeName := startTag[3 : len(startTag)-1]
					fmt.Printf("  %s[%d] name=%q pos=(%s,%s) sz=(%s,%s)%s text=%q\n",
						typeName, count, name, offX, offY, extCx, extCy, extra, text)
					count++
					idx = absStart + endPos + len(endTag)
				}
			}
			return
		}
	}
}
