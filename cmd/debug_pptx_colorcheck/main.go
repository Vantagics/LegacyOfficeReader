package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
)

type SlideXML struct {
	XMLName xml.Name `xml:"sld"`
	CSld    struct {
		SpTree struct {
			Shapes []ShapeXML `xml:"sp"`
		} `xml:"spTree"`
	} `xml:"cSld"`
}

type ShapeXML struct {
	NvSpPr struct {
		CnvPr struct {
			Name string `xml:"name,attr"`
		} `xml:"cNvPr"`
	} `xml:"nvSpPr"`
	SpPr struct {
		SolidFill *struct {
			SrgbClr *struct {
				Val   string `xml:"val,attr"`
				Alpha *struct {
					Val string `xml:"val,attr"`
				} `xml:"alpha"`
			} `xml:"srgbClr"`
		} `xml:"solidFill"`
		NoFill *struct{} `xml:"noFill"`
	} `xml:"spPr"`
	TxBody *struct {
		Paragraphs []ParaXML `xml:"p"`
	} `xml:"txBody"`
}

type ParaXML struct {
	Runs []RunXML `xml:"r"`
}

type RunXML struct {
	RPr *struct {
		SolidFill *struct {
			SrgbClr *struct {
				Val string `xml:"val,attr"`
			} `xml:"srgbClr"`
		} `xml:"solidFill"`
	} `xml:"rPr"`
	Text string `xml:"t"`
}

func isDark(hex string) bool {
	if len(hex) != 6 {
		return false
	}
	hd := func(c byte) int {
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
	r := hd(hex[0])*16 + hd(hex[1])
	g := hd(hex[2])*16 + hd(hex[3])
	b := hd(hex[4])*16 + hd(hex[5])
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	return lum < 128
}

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	issues := 0
	for _, f := range r.File {
		if !strings.HasPrefix(f.Name, "ppt/slides/slide") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		slideNum := strings.TrimPrefix(f.Name, "ppt/slides/slide")
		slideNum = strings.TrimSuffix(slideNum, ".xml")

		rc, err := f.Open()
		if err != nil {
			continue
		}
		data, _ := io.ReadAll(rc)
		rc.Close()

		var slide SlideXML
		if err := xml.Unmarshal(data, &slide); err != nil {
			continue
		}

		for _, shape := range slide.CSld.SpTree.Shapes {
			fillColor := ""
			hasFill := false
			fillIsTransparent := false
			if shape.SpPr.SolidFill != nil && shape.SpPr.SolidFill.SrgbClr != nil {
				fillColor = shape.SpPr.SolidFill.SrgbClr.Val
				hasFill = true
				// Check alpha - if very low, fill is effectively transparent
				if shape.SpPr.SolidFill.SrgbClr.Alpha != nil {
					alphaVal, _ := strconv.Atoi(shape.SpPr.SolidFill.SrgbClr.Alpha.Val)
					if alphaVal < 20000 { // < 20% opacity
						fillIsTransparent = true
					}
				}
			}
			if shape.SpPr.NoFill != nil {
				hasFill = false
			}

			if shape.TxBody == nil {
				continue
			}

			for _, para := range shape.TxBody.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t == "" || run.RPr == nil {
						continue
					}
					textColor := ""
					if run.RPr.SolidFill != nil && run.RPr.SolidFill.SrgbClr != nil {
						textColor = run.RPr.SolidFill.SrgbClr.Val
					}
					if textColor == "" {
						continue
					}

					if hasFill && fillColor != "" && !fillIsTransparent {
						fillDark := isDark(fillColor)
						textDark := isDark(textColor)
						if fillDark && textDark {
							if len(t) > 30 {
								t = t[:30] + "..."
							}
							fmt.Printf("DARK_ON_DARK slide=%s fill=%s text=%s: %q\n", slideNum, fillColor, textColor, t)
							issues++
						} else if !fillDark && !textDark && textColor == fillColor {
							if len(t) > 30 {
								t = t[:30] + "..."
							}
							fmt.Printf("SAME_COLOR slide=%s fill=%s text=%s: %q\n", slideNum, fillColor, textColor, t)
							issues++
						}
					}
				}
			}
		}
	}
	fmt.Printf("\nTotal issues: %d\n", issues)
}
