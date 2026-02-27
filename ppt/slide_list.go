package ppt

import (
	"encoding/binary"
	"fmt"

	"github.com/shakinm/xlsReader/helpers"
)

// parseSlideListWithText scans the PowerPoint Document stream for
// SlideListWithText containers (recType=0x0FF0) and extracts slides
// with their text content.
//
// There are two approaches depending on the PPT file structure:
//  1. Text atoms are directly inside the SlideListWithText container
//     (interleaved with SlidePersistAtom records)
//  2. Text atoms are inside individual slide records referenced by
//     the PersistDirectory (the SlideListWithText only has SlidePersistAtom entries)
//
// This function handles both cases.
func parseSlideListWithText(pptDocData []byte) ([]Slide, error) {
	return parseSlideListWithTextAndPersist(pptDocData, nil)
}

// parseSlideListWithTextAndPersist is the full version that also accepts
// a persist directory for resolving slide references.
func parseSlideListWithTextAndPersist(pptDocData []byte, persistDir map[uint32]uint32) ([]Slide, error) {
	var slides []Slide
	dataLen := uint32(len(pptDocData))
	offset := uint32(0)

	for offset+recordHeaderSize <= dataLen {
		rh, err := readRecordHeader(pptDocData, offset)
		if err != nil {
			break
		}

		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen

		if recDataEnd > dataLen {
			break
		}

		if rh.recType == rtSlideListWithText {
			// [MS-PPT] 2.4.14.1: instance 0 = slides, 1 = master slides, 2 = notes slides
			// Only process instance 0 (actual presentation slides)
			if rh.recInstance() == 0 {
				containerSlides, err := parseSlideListContainer(pptDocData, recDataStart, recDataEnd, persistDir)
				if err != nil {
					return nil, fmt.Errorf("failed to parse SlideListWithText at offset %d: %w", offset, err)
				}
				slides = append(slides, containerSlides...)
			}
		}

		// For container records (recVer == 0xF), step into the container
		// to find nested SlideListWithText records
		if rh.recVer() == 0xF {
			offset = recDataStart // step into container
		} else {
			offset = recDataEnd // skip past atom
		}
	}

	return slides, nil
}

// parseSlideListContainer iterates through sub-records inside a
// SlideListWithText container and groups text atoms by slide.
//
// If the container has inline text atoms, they are used directly.
// If the container only has SlidePersistAtom entries and a persistDir
// is available, text is extracted from the referenced slide records.
func parseSlideListContainer(data []byte, start, end uint32, persistDir map[uint32]uint32) ([]Slide, error) {
	var slides []Slide
	var currentSlide *Slide
	hasInlineText := false
	offset := start

	// Collect SlidePersistAtom persist IDs for fallback
	type slidePersistInfo struct {
		psrReference uint32
	}
	var persistInfos []slidePersistInfo

	for offset+recordHeaderSize <= end {
		rh, err := readRecordHeader(data, offset)
		if err != nil {
			break
		}

		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen

		if recDataEnd > end {
			break
		}

		switch rh.recType {
		case rtSlidePersistAtom:
			if currentSlide != nil {
				slides = append(slides, *currentSlide)
			}
			currentSlide = &Slide{}
			// Extract psrReference (first 4 bytes of SlidePersistAtom body)
			if rh.recLen >= 4 {
				psrRef := binary.LittleEndian.Uint32(data[recDataStart : recDataStart+4])
				persistInfos = append(persistInfos, slidePersistInfo{psrReference: psrRef})
			}

		case rtTextCharsAtom:
			if currentSlide != nil && rh.recLen > 0 {
				text := helpers.DecodeUTF16LE(data[recDataStart:recDataEnd])
				if text != "" {
					currentSlide.texts = append(currentSlide.texts, text)
					hasInlineText = true
				}
			}

		case rtTextBytesAtom:
			if currentSlide != nil && rh.recLen > 0 {
				text := helpers.DecodeANSI(data[recDataStart:recDataEnd])
				if text != "" {
					currentSlide.texts = append(currentSlide.texts, text)
					hasInlineText = true
				}
			}
		}

		offset = recDataEnd
	}

	// Don't forget the last slide
	if currentSlide != nil {
		slides = append(slides, *currentSlide)
	}

	// If no inline text was found and we have a persist directory,
	// try to extract text from individual slide records
	if !hasInlineText && persistDir != nil && len(persistInfos) > 0 {
		slides = nil // Reset slides, rebuild from persist directory
		for _, info := range persistInfos {
			slideOffset, ok := persistDir[info.psrReference]
			if !ok {
				slides = append(slides, Slide{})
				continue
			}
			slide := extractSlideTexts(data, slideOffset)
			slides = append(slides, slide)
		}
	}

	return slides, nil
}

// extractSlideTexts recursively scans a slide record for text atoms
// and returns a Slide with all found text.
func extractSlideTexts(data []byte, slideOffset uint32) Slide {
	slide := Slide{}
	if uint32(len(data)) <= slideOffset+recordHeaderSize {
		return slide
	}

	rh, err := readRecordHeader(data, slideOffset)
	if err != nil {
		return slide
	}

	recEnd := slideOffset + recordHeaderSize + rh.recLen
	if recEnd > uint32(len(data)) {
		return slide
	}

	// Recursively scan for text atoms
	collectTexts(data, slideOffset+recordHeaderSize, recEnd, &slide)
	return slide
}

// collectTexts recursively scans records for text atoms.
func collectTexts(data []byte, start, end uint32, slide *Slide) {
	offset := start
	for offset+recordHeaderSize <= end {
		rh, err := readRecordHeader(data, offset)
		if err != nil {
			break
		}

		recDataStart := offset + recordHeaderSize
		recDataEnd := recDataStart + rh.recLen

		if recDataEnd > end {
			break
		}

		switch rh.recType {
		case rtTextCharsAtom:
			if rh.recLen > 0 {
				text := helpers.DecodeUTF16LE(data[recDataStart:recDataEnd])
				if text != "" {
					slide.texts = append(slide.texts, text)
				}
			}
		case rtTextBytesAtom:
			if rh.recLen > 0 {
				text := helpers.DecodeANSI(data[recDataStart:recDataEnd])
				if text != "" {
					slide.texts = append(slide.texts, text)
				}
			}
		}

		// If it's a container, recurse into it
		if rh.recVer() == 0xF {
			collectTexts(data, recDataStart, recDataEnd, slide)
			offset = recDataEnd
		} else {
			offset = recDataEnd
		}
	}
}
