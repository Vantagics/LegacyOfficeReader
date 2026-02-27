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
		if f.Name != "ppt/slides/slide41.xml" {
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
			if sp.TxBody == nil {
				continue
			}
			for _, p := range sp.TxBody.Paras {
				for _, r := range p.Runs {
					t := strings.TrimSpace(r.Text)
					if t == "" {
						continue
					}
					tc := ""
					if r.RPr != nil && r.RPr.SolidFill != nil && r.RPr.SolidFill.SrgbClr != nil {
						tc = r.RPr.SolidFill.SrgbClr.Val
					}
					if len(t) > 30 {
						t = t[:30] + "..."
					}
					fmt.Printf("id=%s fill=%s textColor=%s: %q\n", sp.NvSpPr.CnvPr.ID, fill, tc, t)
				}
			}
		}
	}
}
