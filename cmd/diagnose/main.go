package main

import (
"fmt"
"strings"

"github.com/shakinm/xlsReader/ppt"
)

func main() {
p, err := ppt.OpenFile("testfie/test.ppt")
if err != nil {
fmt.Println("ERROR:", err)
return
}
slides := p.GetSlides()
masters := p.GetMasters()
images := p.GetImages()

fmt.Printf("Slides: %d, Masters: %d, Images: %d\n", len(slides), len(masters), len(images))

// Analyze master shapes per layout
masterRefToIdx := make(map[uint32]int)
var layoutRefs []uint32
for _, s := range slides {
ref := s.GetMasterRef()
if _, ok := masterRefToIdx[ref]; !ok {
masterRefToIdx[ref] = len(layoutRefs)
layoutRefs = append(layoutRefs, ref)
}
}

fmt.Printf("\n=== Layout Analysis (%d unique layouts) ===\n", len(layoutRefs))
for li, ref := range layoutRefs {
m, ok := masters[ref]
if !ok {
fmt.Printf("Layout %d (ref=%d): NO MASTER DATA\n", li+1, ref)
continue
}

// Count slides using this layout
slideCount := 0
for _, s := range slides {
if s.GetMasterRef() == ref {
slideCount++
}
}

totalShapes := len(m.Shapes)
textShapes := 0
imageShapes := 0
placeholderShapes := 0
decorativeShapes := 0

for _, sh := range m.Shapes {
if sh.IsImage {
imageShapes++
}
if sh.IsText && len(sh.Paragraphs) > 0 {
textShapes++
isPlaceholder := false
for _, para := range sh.Paragraphs {
for _, run := range para.Runs {
t := strings.TrimSpace(run.Text)
if strings.Contains(t, "编辑母版") || strings.Contains(t, "Click to edit") {
isPlaceholder = true
}
}
}
if isPlaceholder {
placeholderShapes++
}
}
if !sh.IsText && !sh.IsImage {
decorativeShapes++
}
}

fmt.Printf("Layout %d (ref=%d): %d slides, %d shapes (text=%d, img=%d, deco=%d, placeholder=%d)\n",
li+1, ref, slideCount, totalShapes, textShapes, imageShapes, decorativeShapes, placeholderShapes)
fmt.Printf("  Background: has=%v, color=%q, imgIdx=%d\n",
m.Background.HasBackground, m.Background.FillColor, m.Background.ImageIdx)

// Show first few shapes
for si, sh := range m.Shapes {
if si >= 5 {
fmt.Printf("  ... and %d more shapes\n", len(m.Shapes)-5)
break
}
text := ""
if sh.IsText {
for _, para := range sh.Paragraphs {
for _, run := range para.Runs {
text += run.Text
}
}
if len(text) > 60 {
text = text[:60] + "..."
}
}
fmt.Printf("  Shape %d: type=%d, pos=(%d,%d), size=(%d,%d), text=%v, img=%v(idx=%d), fill=%q, text=%q\n",
si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsText, sh.IsImage, sh.ImageIdx, sh.FillColor, text)
}
}

// Check first few slides for their content
fmt.Printf("\n=== First 5 Slides Detail ===\n")
for i := 0; i < 5 && i < len(slides); i++ {
s := slides[i]
shapes := s.GetShapes()
bg := s.GetBackground()
fmt.Printf("Slide %d: %d shapes, masterRef=%d, bg=%v\n", i+1, len(shapes), s.GetMasterRef(), bg.HasBackground)

textCount := 0
imgCount := 0
for _, sh := range shapes {
if sh.IsText {
textCount++
}
if sh.IsImage {
imgCount++
}
}
fmt.Printf("  text=%d, img=%d\n", textCount, imgCount)

// Show first few shapes
for si, sh := range shapes {
if si >= 8 {
fmt.Printf("  ... and %d more shapes\n", len(shapes)-8)
break
}
text := ""
if sh.IsText {
for _, para := range sh.Paragraphs {
for _, run := range para.Runs {
text += run.Text
}
}
if len(text) > 50 {
text = text[:50] + "..."
}
}
fmt.Printf("  sh%d: type=%d, (%d,%d) %dx%d, text=%v, img=%v(idx=%d), fill=%q, noFill=%v, text=%q\n",
si, sh.ShapeType, sh.Left, sh.Top, sh.Width, sh.Height, sh.IsText, sh.IsImage, sh.ImageIdx, sh.FillColor, sh.NoFill, text)
}
}
}