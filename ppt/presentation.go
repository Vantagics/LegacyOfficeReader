package ppt

import (
	"errors"

	"github.com/shakinm/xlsReader/common"
)

// Presentation represents a parsed PPT file.
type Presentation struct {
	slides      []Slide
	images      []common.Image
	fonts       []string
	slideWidth  int32
	slideHeight int32
	masters     map[uint32]MasterSlide
}

// GetSlides returns all slides in the presentation.
func (p *Presentation) GetSlides() []Slide {
	return p.slides
}

// GetSlide returns the slide at the given index.
func (p *Presentation) GetSlide(index int) (*Slide, error) {
	if index < 0 || index >= len(p.slides) {
		return nil, errors.New("slide index out of range")
	}
	return &p.slides[index], nil
}

// GetNumberSlides returns the total number of slides.
func (p *Presentation) GetNumberSlides() int {
	return len(p.slides)
}

// GetImages returns all embedded images extracted from the presentation.
// It always returns a non-nil slice; if there are no images, it returns an empty slice.
func (p *Presentation) GetImages() []common.Image {
	if p.images != nil {
		return p.images
	}
	return []common.Image{}
}

// GetFonts returns the font name index table.
func (p *Presentation) GetFonts() []string {
	if p.fonts != nil {
		return p.fonts
	}
	return []string{}
}

// GetSlideSize returns the slide width and height in EMU.
func (p *Presentation) GetSlideSize() (int32, int32) {
	if p.slideWidth == 0 && p.slideHeight == 0 {
		return 9144000, 6858000 // default 10" x 7.5"
	}
	return p.slideWidth, p.slideHeight
}

// GetMasters returns the parsed slide masters keyed by persist reference.
func (p *Presentation) GetMasters() map[uint32]MasterSlide {
	if p.masters != nil {
		return p.masters
	}
	return map[uint32]MasterSlide{}
}
