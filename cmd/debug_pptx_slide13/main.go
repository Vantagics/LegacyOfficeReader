package main

import (
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
)

type Sld struct {
	XMLName xml.Name `xml:"sld"`
	CSld    struct {
		SpTree struct {
			Shapes []Sp `xml:"sp"`
		} `xml:"spTree"`
	} `xml:"cSld"`
}

type Sp struct {
	NvSpPr struct {
		CnvPr struct {
			ID   string `xml:"id,attr"`
			Name string `xml:"name,attr"`
		} `xml:"cNvPr"`
	} `xml:"nvSpPr"`
	SpPr struct {
		Xfrm *struct {
			Off struct {
				X string `xml:"x,attr"`
				Y string `xml:"y,attr"`
			} `xml:"off"`
			Ext struct {
				CX string `xml:"cx,attr"`
				CY string `xml:"cy,attr"`
			} `xml:"ext"`
		} `xml:"xfrm"`
		SolidFill *struct {
			SrgbClr *struct {
				Val string `xml:"val,attr"`
			} `xml:"srgbClr"`
		} `xml:"solidFill"`
		NoFill *struct{} `xml:"noFill"`
	} `xml:"spPr"`
	TxBody *struct {
		Paras []Para `xml:"p"`
	} `xml:"txBody"`
}

type Para struct {
	Runs []Run `xml:"r"`
}

type Run struct {
	RPr *struct {
		Sz        string `xml:"sz,attr"`
		SolidFill *struct {
			SrgbClr *struct {
				Val string `xml:"val,attr"`
			} `xml:"srgbClr"`
		} `xml:"solidFill"`
	} `xml:"rPr"`
	Text string `xml:"t"`
}

func main() {
	r, err := zip.OpenReader("testfie/test.pptx")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name != "ppt/slides/slide13.xml" {
			continue
		}
		rc, _ := f.Open()
		data, _ := io.ReadAll(rc)
		rc.Close()

		var sld Sld
		xml.Unmarshal(data, &sld)

		for _, sp := range sld.CSld.SpTree.Shapes {
			fill := ""
			if sp.SpPr.SolidFill != nil && sp.SpPr.SolidFill.SrgbClr != nil {
				fill = sp.SpPr.SolidFill.SrgbClr.Val
			}
			if sp.SpPr.NoFill != nil {
				fill = "noFill"
			}
			pos := ""
			if sp.SpPr.Xfrm != nil {
				pos = fmt.Sprintf("(%s,%s) %sx%s", sp.SpPr.Xfrm.Off.X, sp.SpPr.Xfrm.Off.Y, sp.SpPr.Xfrm.Ext.CX, sp.SpPr.Xfrm.Ext.CY)
			}
			hasText := false
			if sp.TxBody != nil {
				for _, p := range sp.TxBody.Paras {
					for _, r := range p.Runs {
						if strings.TrimSpace(r.Text) != "" {
							hasText = true
						}
					}
				}
			}
			if !hasText {
				continue
			}
			fmt.Printf("Shape id=%s fill=%s pos=%s\n", sp.NvSpPr.CnvPr.ID, fill, pos)
			if sp.TxBody != nil {
				for pi, p := range sp.TxBody.Paras {
					for ri, r := range p.Runs {
						tc := ""
						if r.RPr != nil && r.RPr.SolidFill != nil && r.RPr.SolidFill.SrgbClr != nil {
							tc = r.RPr.SolidFill.SrgbClr.Val
						}
						sz := ""
						if r.RPr != nil {
							sz = r.RPr.Sz
						}
						t := r.Text
						if len(t) > 40 {
							t = t[:40] + "..."
						}
						fmt.Printf("  P[%d]R[%d] color=%s sz=%s: %q\n", pi, ri, tc, sz, t)
					}
				}
			}
		}
	}
}
