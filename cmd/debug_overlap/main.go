package main

import (
	"fmt"
	"os"

	"github.com/shakinm/xlsReader/ppt"
)

func main() {
	pres, err := ppt.OpenFile("testfie/test.ppt")
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed: %v\n", err)
		os.Exit(1)
	}

	slides := pres.GetSlides()
	masters := pres.GetMasters()

	// Layout 4 (ref=2147483734) - the most common layout with 56 slides
	// Has watermark image at imgIdx=13 and logo at imgIdx=11
	ref := uint32(2147483734)
	m, ok := masters[ref]
	if !ok {
		fmt.Println("Master not found")
		return
	}

	fmt.Println("=== Layout 4 (ref=2147483734) shapes ===")
	for i, sh := range m.Shapes {
		fmt.Printf("Shape[%d]: type=%d pos=(%d,%d) size=(%d,%d) isImg=%v imgIdx=%d\n",
			i, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsImage, sh.ImageIdx)
	}

	// The watermark is imgIdx=13: pos=(6026150,4087812) size=(6165850,3016250)
	// So it covers x=6026150 to 12192000, y=4087812 to 7104062
	// The logo is imgIdx=11: pos=(10085387,-395287) size=(1884362,1882775)
	// So it covers x=10085387 to 11969749, y=-395287 to 1487488

	// Check which "invisible" shapes actually overlap with these images
	fmt.Println("\n=== Checking invisible text shapes on layout 4 slides ===")
	checkSlides := []int{9, 11, 12, 27, 41, 61}
	for _, sn := range checkSlides {
		if sn > len(slides) {
			continue
		}
		s := slides[sn-1]
		if s.GetMasterRef() != ref {
			fmt.Printf("Slide %d: NOT on layout 4 (ref=%d)\n", sn, s.GetMasterRef())
			continue
		}
		shapes := s.GetShapes()
		fmt.Printf("\nSlide %d: %d shapes\n", sn, len(shapes))
		for si, sh := range shapes {
			if len(sh.Paragraphs) == 0 {
				continue
			}
			for _, p := range sh.Paragraphs {
				for _, r := range p.Runs {
					if r.Color != "FFFFFF" {
						continue
					}
					text := r.Text
					if len([]rune(text)) > 30 {
						text = string([]rune(text)[:30]) + "..."
					}
					// Check overlap with watermark (imgIdx=13)
					wmLeft := int64(6026150)
					wmTop := int64(4087812)
					wmRight := wmLeft + int64(6165850)
					wmBottom := wmTop + int64(3016250)
					cx := int64(sh.Left) + int64(sh.Width)/2
					cy := int64(sh.Top) + int64(sh.Height)/2
					overlapWM := cx >= wmLeft && cx <= wmRight && cy >= wmTop && cy <= wmBottom

					// Check overlap with logo (imgIdx=11)
					logoLeft := int64(10085387)
					logoTop := int64(-395287)
					logoRight := logoLeft + int64(1884362)
					logoBottom := logoTop + int64(1882775)
					overlapLogo := cx >= logoLeft && cx <= logoRight && cy >= logoTop && cy <= logoBottom

					// Check if in title area (above connector at y=1212850)
					inTitleArea := cy < 1212850

					fmt.Printf("  Shape[%d] pos=(%d,%d) sz=(%d,%d) center=(%d,%d) fill=%s noFill=%v\n",
						si, sh.Left, sh.Top, sh.Width, sh.Height, cx, cy, sh.FillColor, sh.NoFill)
					fmt.Printf("    Text: %s  raw=0x%08X\n", text, r.ColorRaw)
					fmt.Printf("    overlapWM=%v overlapLogo=%v inTitle=%v\n", overlapWM, overlapLogo, inTitleArea)
					goto nextShape
				}
			}
		nextShape:
		}
	}
}
