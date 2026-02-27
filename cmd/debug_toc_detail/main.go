package main

import (
	"fmt"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, _ := ppt.OpenFile("testfie/test.ppt")

	// Slide 2 (index 1)
	slides := p.GetSlides()
	slide := slides[1]
	shapes := slide.GetShapes()
	fmt.Printf("Slide 2: %d shapes\n", len(shapes))
	for i, sh := range shapes {
		text := ""
		for _, para := range sh.Paragraphs {
			for _, run := range para.Runs {
				text += run.Text
			}
		}
		fmt.Printf("  Shape %d: type=%d pos=(%d,%d) size=(%d,%d)\n", i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height)
		fmt.Printf("    fill='%s' noFill=%v line='%s' noLine=%v lineW=%d lineDash=%d\n",
			sh.FillColor, sh.NoFill, sh.LineColor, sh.NoLine, sh.LineWidth, sh.LineDash)
		fmt.Printf("    geoVerts=%d geoSegs=%d geoRect=(%d,%d,%d,%d)\n",
			len(sh.GeoVertices), len(sh.GeoSegments), sh.GeoLeft, sh.GeoTop, sh.GeoRight, sh.GeoBottom)
		if len(sh.GeoVertices) > 0 {
			fmt.Printf("    vertices: ")
			for _, v := range sh.GeoVertices {
				fmt.Printf("(%d,%d) ", v.X, v.Y)
			}
			fmt.Println()
			fmt.Printf("    segments: ")
			for _, s := range sh.GeoSegments {
				fmt.Printf("(type=%d,cnt=%d) ", s.SegType, s.Count)
			}
			fmt.Println()
		}
		fmt.Printf("    text='%s'\n", text)
	}
}
