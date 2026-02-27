package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"

	"github.com/shakinm/xlsReader/convert/pptconv"
	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	// Step 1: Parse PPT source
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPT: %v\n", err)
		os.Exit(1)
	}

	// Step 2: Convert to PPTX
	outFile := "testfie/test_compare.pptx"
	if err := pptconv.ConvertFile("testfie/test.ppt", outFile); err != nil {
		fmt.Fprintf(os.Stderr, "Error converting: %v\n", err)
		os.Exit(1)
	}

	// Step 3: Compare each slide
	slides := p.GetSlides()
	fmt.Printf("Total PPT slides: %d\n\n", len(slides))

	// Open the converted PPTX
	zr, err := zip.OpenReader(outFile)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error opening PPTX: %v\n", err)
		os.Exit(1)
	}
	defer zr.Close()

	for i, slide := range slides {
		slideNum := i + 1
		fmt.Printf("========== Slide %d ==========\n", slideNum)

		// PPT side
		shapes := slide.GetShapes()
		fmt.Printf("[PPT] Shapes: %d\n", len(shapes))
		for si, sh := range shapes {
			if len(sh.Paragraphs) == 0 && !sh.IsImage {
				continue
			}
			fmt.Printf("  PPT Shape[%d]: type=%d pos=(%d,%d) size=(%d,%d) fill=%s isImg=%v\n",
				si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.FillColor, sh.IsImage)
			for pi, para := range sh.Paragraphs {
				for ri, run := range para.Runs {
					text := strings.TrimSpace(run.Text)
					if text == "" {
						continue
					}
					if len(text) > 60 {
						text = text[:60] + "..."
					}
					fmt.Printf("    P%d/R%d: size=%d color=%s bold=%v font=%q text=%q\n",
						pi, ri, run.FontSize, run.Color, run.Bold, run.FontName, text)
				}
			}
		}

		// PPTX side
		slideFile := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		pptxData := readZipFile(zr, slideFile)
		if pptxData == nil {
			fmt.Printf("[PPTX] slide%d.xml not found\n", slideNum)
			continue
		}
		pptxShapes := parsePptxSlide(pptxData)
		fmt.Printf("[PPTX] Shapes: %d\n", len(pptxShapes))
		for si, sh := range pptxShapes {
			if len(sh.runs) == 0 && !sh.isPic {
				continue
			}
			fmt.Printf("  PPTX Shape[%d]: pos=(%d,%d) size=(%d,%d) fill=%s isPic=%v\n",
				si, sh.left, sh.top, sh.width, sh.height, sh.fillColor, sh.isPic)
			for ri, run := range sh.runs {
				text := strings.TrimSpace(run.text)
				if text == "" {
					continue
				}
				if len(text) > 60 {
					text = text[:60] + "..."
				}
				fmt.Printf("    R%d: sz=%d color=%s bold=%v font=%q text=%q\n",
					ri, run.fontSize, run.color, run.bold, run.fontName, text)
			}
		}
		fmt.Println()
	}
}

func readZipFile(zr *zip.ReadCloser, name string) []byte {
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil
			}
			defer rc.Close()
			data, _ := io.ReadAll(rc)
			return data
		}
	}
	return nil
}

type pptxShape struct {
	left, top, width, height int64
	fillColor                string
	isPic                    bool
	runs                     []pptxRun
}

type pptxRun struct {
	text     string
	fontName string
	fontSize int
	bold     bool
	color    string
}

func parsePptxSlide(data []byte) []pptxShape {
	var shapes []pptxShape
	d := xml.NewDecoder(strings.NewReader(string(data)))

	type state struct {
		inSp     bool
		inPic    bool
		inSpPr   bool
		inTxBody bool
		inRPr    bool
		inR      bool
		inT      bool
		inOff    bool
		inExt    bool
		inFill   bool
	}
	var s state
	var curShape *pptxShape
	var curRun *pptxRun
	var textBuf strings.Builder

	for {
		tok, err := d.Token()
		if err != nil {
			break
		}
		switch t := tok.(type) {
		case xml.StartElement:
			local := t.Name.Local
			switch local {
			case "sp":
				shapes = append(shapes, pptxShape{})
				curShape = &shapes[len(shapes)-1]
				s.inSp = true
			case "pic":
				shapes = append(shapes, pptxShape{isPic: true})
				curShape = &shapes[len(shapes)-1]
				s.inPic = true
			case "spPr":
				s.inSpPr = true
			case "txBody":
				s.inTxBody = true
			case "off":
				if s.inSpPr && curShape != nil {
					for _, a := range t.Attr {
						switch a.Name.Local {
						case "x":
							curShape.left, _ = strconv.ParseInt(a.Value, 10, 64)
						case "y":
							curShape.top, _ = strconv.ParseInt(a.Value, 10, 64)
						}
					}
				}
			case "ext":
				if s.inSpPr && curShape != nil {
					for _, a := range t.Attr {
						switch a.Name.Local {
						case "cx":
							curShape.width, _ = strconv.ParseInt(a.Value, 10, 64)
						case "cy":
							curShape.height, _ = strconv.ParseInt(a.Value, 10, 64)
						}
					}
				}
			case "srgbClr":
				if s.inSpPr && curShape != nil && !s.inTxBody {
					for _, a := range t.Attr {
						if a.Name.Local == "val" {
							curShape.fillColor = a.Value
						}
					}
				}
				if s.inRPr && curRun != nil {
					for _, a := range t.Attr {
						if a.Name.Local == "val" {
							curRun.color = a.Value
						}
					}
				}
			case "r":
				if s.inTxBody {
					s.inR = true
					curRun = &pptxRun{}
				}
			case "rPr":
				if s.inR && curRun != nil {
					s.inRPr = true
					for _, a := range t.Attr {
						switch a.Name.Local {
						case "sz":
							curRun.fontSize, _ = strconv.Atoi(a.Value)
						case "b":
							curRun.bold = a.Value == "1"
						}
					}
				}
			case "latin", "ea":
				if s.inRPr && curRun != nil && curRun.fontName == "" {
					for _, a := range t.Attr {
						if a.Name.Local == "typeface" {
							curRun.fontName = a.Value
						}
					}
				}
			case "t":
				if s.inR {
					s.inT = true
					textBuf.Reset()
				}
			}
		case xml.CharData:
			if s.inT {
				textBuf.Write(t)
			}
		case xml.EndElement:
			local := t.Name.Local
			switch local {
			case "sp":
				s.inSp = false
				curShape = nil
			case "pic":
				s.inPic = false
				curShape = nil
			case "spPr":
				s.inSpPr = false
			case "txBody":
				s.inTxBody = false
			case "r":
				if s.inR && curRun != nil && curShape != nil {
					curRun.text = textBuf.String()
					curShape.runs = append(curShape.runs, *curRun)
				}
				s.inR = false
				curRun = nil
			case "rPr":
				s.inRPr = false
			case "t":
				s.inT = false
			}
		}
	}
	return shapes
}
