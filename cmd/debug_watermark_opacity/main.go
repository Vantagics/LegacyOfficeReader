package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	f, err := os.Open("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}
	defer f.Close()

	p, err := ppt.OpenReader(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Parse error: %v\n", err)
		os.Exit(1)
	}

	masters := p.GetMasters()
	m := masters[uint32(2147483734)]
	for i, sh := range m.Shapes {
		fmt.Printf("shape[%d]: type=%d isImage=%v imageIdx=%d\n", i, sh.ShapeType, sh.IsImage, sh.ImageIdx)
		fmt.Printf("  pos=(%d,%d) size=(%d,%d)\n", sh.Left, sh.Top, sh.Width, sh.Height)
		fmt.Printf("  FillColor=%q NoFill=%v FillOpacity=%d\n", sh.FillColor, sh.NoFill, sh.FillOpacity)
		fmt.Printf("  LineColor=%q NoLine=%v LineWidth=%d\n", sh.LineColor, sh.NoLine, sh.LineWidth)
		fmt.Printf("  Rotation=%d FlipH=%v FlipV=%v\n", sh.Rotation, sh.FlipH, sh.FlipV)
	}
}
