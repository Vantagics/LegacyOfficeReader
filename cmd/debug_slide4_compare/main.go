package main

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// PPT source data
	p, _ := ppt.OpenFile("testfie/test.ppt")
	slides := p.GetSlides()
	slide := slides[3] // slide 4 (0-indexed)
	shapes := slide.GetShapes()
	scheme := slide.GetColorScheme()
	defaults := slide.GetDefaultTextStyles()

	fmt.Printf("Slide 4: %d shapes, scheme=%v\n", len(shapes), scheme)
	fmt.Printf("Master defaults:\n")
	for i, d := range defaults {
		fmt.Printf("  Level %d: size=%d font='%s' color='%s' colorRaw=0x%08X bold=%v\n",
			i, d.FontSize, d.FontName, d.Color, d.ColorRaw, d.Bold)
	}
	fmt.Println()

	for i, sh := range shapes {
		text := ""
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				text += run.Text
			}
		}
		if len(text) > 80 {
			text = text[:80] + "..."
		}
		fmt.Printf("Shape %d: type=%d pos=(%d,%d) size=(%d,%d) text='%s'\n",
			i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, text)
		fmt.Printf("  fill='%s' noFill=%v line='%s' noLine=%v\n",
			sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine)

		for j, para := range sh.Paragraphs {
			if j > 5 {
				fmt.Printf("  ... (%d more paragraphs)\n", len(sh.Paragraphs)-j)
				break
			}
			paraText := ""
			for _, r := range para.Runs {
				paraText += r.Text
			}
			if len(paraText) > 60 {
				paraText = paraText[:60] + "..."
			}
			fmt.Printf("  Para %d: align=%d indent=%d bullet=%v text='%s'\n",
				j, para.Alignment, para.IndentLevel, para.HasBullet, paraText)
			for _, r := range para.Runs {
				rt := r.Text
				if len(rt) > 40 {
					rt = rt[:40] + "..."
				}
				fmt.Printf("    Run: sz=%d font='%s' color='%s' raw=0x%08X bold=%v italic=%v text='%s'\n",
					r.FontSize, r.FontName, r.Color, r.ColorRaw, r.Bold, r.Italic, rt)
			}
		}
		fmt.Println()
	}

	// Now check PPTX output
	fmt.Println("\n=== PPTX OUTPUT ===")
	f, _ := zip.OpenReader("testfie/test.pptx")
	defer f.Close()
	for _, zf := range f.File {
		if zf.Name == "ppt/slides/slide4.xml" {
			rc, _ := zf.Open()
			data, _ := io.ReadAll(rc)
			rc.Close()
			xml := string(data)

			// Find all shapes and extract font sizes and colors
			idx := 0
			spNum := 0
			for {
				// Look for both sp and pic
				spStart := strings.Index(xml[idx:], "<p:sp>")
				if spStart < 0 {
					break
				}
				spStart += idx
				spEnd := strings.Index(xml[spStart:], "</p:sp>")
				if spEnd < 0 {
					break
				}
				spEnd += spStart + len("</p:sp>")
				snippet := xml[spStart:spEnd]
				spNum++

				// Extract text
				textParts := []string{}
				tIdx := 0
				for {
					tStart := strings.Index(snippet[tIdx:], "<a:t>")
					if tStart < 0 {
						break
					}
					tStart += tIdx + 5
					tEnd := strings.Index(snippet[tStart:], "</a:t>")
					if tEnd < 0 {
						break
					}
					textParts = append(textParts, snippet[tStart:tStart+tEnd])
					tIdx = tStart + tEnd
				}
				text := strings.Join(textParts, " | ")
				if len(text) > 80 {
					text = text[:80] + "..."
				}

				// Extract font sizes (sz="...")
				sizes := []string{}
				sIdx := 0
				for {
					szStart := strings.Index(snippet[sIdx:], ` sz="`)
					if szStart < 0 {
						break
					}
					szStart += sIdx + 5
					szEnd := strings.Index(snippet[szStart:], `"`)
					if szEnd < 0 {
						break
					}
					sizes = append(sizes, snippet[szStart:szStart+szEnd])
					sIdx = szStart + szEnd
				}

				// Extract colors (val="XXXXXX" in srgbClr)
				colors := []string{}
				cIdx := 0
				for {
					cStart := strings.Index(snippet[cIdx:], `<a:srgbClr val="`)
					if cStart < 0 {
						break
					}
					cStart += cIdx + 16
					cEnd := strings.Index(snippet[cStart:], `"`)
					if cEnd >= 0 {
						colors = append(colors, snippet[cStart:cStart+cEnd])
					}
					cIdx = cStart + 6
				}

				fmt.Printf("PPTX Shape %d: sizes=%v colors=%v text='%s'\n", spNum, sizes, colors, text)
				idx = spEnd
			}
		}
	}
}
