package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	p, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	masters := p.GetMasters()
	slides := p.GetSlides()

	// Find which master refs are used by slides
	usedMasters := map[uint32]bool{}
	for _, s := range slides {
		usedMasters[s.GetMasterRef()] = true
	}

	for ref, m := range masters {
		if !usedMasters[ref] {
			continue
		}
		fmt.Printf("Master ref=%d: shapes=%d bg=(has=%v,color=%s,img=%d)\n",
			ref, len(m.Shapes), m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)

		for i, sh := range m.Shapes {
			var texts []string
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t != "" {
						texts = append(texts, t)
					}
				}
			}

			isPlaceholder := false
			for _, para := range sh.Paragraphs {
				for _, run := range para.Runs {
					t := strings.TrimSpace(run.Text)
					if t == "" {
						continue
					}
					prefixes := []string{"单击此处编辑母版", "点击此处编辑母版", "Click to edit Master"}
					for _, prefix := range prefixes {
						if strings.Contains(t, prefix) {
							isPlaceholder = true
						}
					}
				}
			}

			w, h := p.GetSlideSize()
			isFullPage := sh.IsImage && sh.ImageIdx >= 0 && sh.Width > int32(float64(w)*0.7) && sh.Height > int32(float64(h)*0.7)

			fmt.Printf("  Shape[%d]: type=%d isText=%v isImage=%v imgIdx=%d pos=(%d,%d) size=(%d,%d)\n",
				i, sh.ShapeType, sh.IsText, sh.IsImage, sh.ImageIdx, sh.Left, sh.Top, sh.Width, sh.Height)
			fmt.Printf("    isPlaceholder=%v isFullPage=%v texts=%v\n", isPlaceholder, isFullPage, texts)
		}
	}
}
